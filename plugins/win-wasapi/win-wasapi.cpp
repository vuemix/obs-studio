#include "enum-wasapi.hpp"

#include <obs-module.h>
#include <obs.h>
#include <util/platform.h>
#include <util/windows/HRError.hpp>
#include <util/windows/ComPtr.hpp>
#include <util/windows/WinHandle.hpp>
#include <util/windows/CoTaskMemPtr.hpp>
#include <util/threading.h>

#include <deque>

using namespace std;

#define OPT_DEVICE_ID         "device_id"
#define OPT_USE_DEVICE_TIMING "use_device_timing"
#define OPT_DISABLE_AEC       "disable_echo_cancellation"
#define OPT_IN_FMT_MODE       "input_format_mode"
#define OPT_AEC_IN_DELAY      "aec_input_delay"
#define OPT_AEC_DUMP_DIR      "aec_dump_file_dir"

static void GetWASAPIDefaults(obs_data_t *settings);

#define OBS_KSAUDIO_SPEAKER_4POINT1 \
		(KSAUDIO_SPEAKER_SURROUND|SPEAKER_LOW_FREQUENCY)

class CMediaBuffer : public IMediaBuffer
{
private:
	DWORD        m_cbLength;
	const DWORD  m_cbMaxLength;
	LONG         m_nRefCount;  // Reference count
	BYTE         *m_pbData;
	UINT64       m_timestamp; // in 100-nanosecond units

	CMediaBuffer(DWORD cbMaxLength, UINT64 timestamp, HRESULT& hr) :
		m_nRefCount(1),
		m_cbMaxLength(cbMaxLength),
		m_cbLength(0),
		m_pbData(NULL),
		m_timestamp(timestamp)
	{
		m_pbData = new BYTE[cbMaxLength];
		if (!m_pbData)
		{
			hr = E_OUTOFMEMORY;
		}
	}

	~CMediaBuffer()
	{
		if (m_pbData)
		{
			delete[] m_pbData;
		}
	}

public:

	// Function to create a new IMediaBuffer object and return
	// an AddRef'd interface pointer.
	static HRESULT Create(long cbMaxLen, UINT64 timestamp, IMediaBuffer **ppBuffer)
	{
		HRESULT hr = S_OK;
		CMediaBuffer *pBuffer = NULL;

		if (ppBuffer == NULL)
		{
			return E_POINTER;
		}

		pBuffer = new CMediaBuffer(cbMaxLen, timestamp, hr);

		if (pBuffer == NULL)
		{
			hr = E_OUTOFMEMORY;
		}

		if (SUCCEEDED(hr))
		{
			*ppBuffer = pBuffer;
			(*ppBuffer)->AddRef();
		}

		if (pBuffer)
		{
			pBuffer->Release();
		}
		return hr;
	}

	UINT64 GetTimestamp()
	{
		return m_timestamp;
	}

	void Dump(FILE* dumpFile)
	{
		fwrite(m_pbData, 1, m_cbLength, dumpFile);
	}

	// IUnknown methods.
	STDMETHODIMP QueryInterface(REFIID riid, void **ppv)
	{
		if (ppv == NULL)
		{
			return E_POINTER;
		}
		else if (riid == __uuidof(IMediaBuffer) || riid == IID_IUnknown)
		{
			*ppv = static_cast<IMediaBuffer *>(this);
			AddRef();
			return S_OK;
		}
		else
		{
			*ppv = NULL;
			return E_NOINTERFACE;
		}
	}

	STDMETHODIMP_(ULONG) AddRef()
	{
		return InterlockedIncrement(&m_nRefCount);
	}

	STDMETHODIMP_(ULONG) Release()
	{
		LONG lRef = InterlockedDecrement(&m_nRefCount);
		if (lRef == 0)
		{
			delete this;
			// m_cRef is no longer valid! Return lRef.
		}
		return lRef;
	}

	// IMediaBuffer methods.
	STDMETHODIMP SetLength(DWORD cbLength)
	{
		if (cbLength > m_cbMaxLength)
		{
			return E_INVALIDARG;
		}
		m_cbLength = cbLength;
		return S_OK;
	}

	STDMETHODIMP GetMaxLength(DWORD *pcbMaxLength)
	{
		if (pcbMaxLength == NULL)
		{
			return E_POINTER;
		}
		*pcbMaxLength = m_cbMaxLength;
		return S_OK;
	}

	STDMETHODIMP GetBufferAndLength(BYTE **ppbBuffer, DWORD *pcbLength)
	{
		// Either parameter can be NULL, but not both.
		if (ppbBuffer == NULL && pcbLength == NULL)
		{
			return E_POINTER;
		}
		if (ppbBuffer)
		{
			*ppbBuffer = m_pbData;
		}
		if (pcbLength)
		{
			*pcbLength = m_cbLength;
		}
		return S_OK;
	}
};

class WASAPISource {
	ComPtr<IMMDevice>           device;
	ComPtr<IMMDevice>           deviceRender;
	ComPtr<IAudioClient>        client;
	ComPtr<IAudioClient>        clientRender;
	ComPtr<IAudioCaptureClient> capture;
	ComPtr<IAudioCaptureClient> captureRender;
	ComPtr<IAudioRenderClient>  render;
	ComPtr<IMediaObject>        captureDMO;
	ComPtr<IMediaBuffer>        captureDMOBuffer;
	CoTaskMemPtr<WAVEFORMATEX>  wfexClient;
	CoTaskMemPtr<WAVEFORMATEX>  wfexClientRender;

	bool                        disableAEC = false;
	int                         inFormatMode = 0;
	int                         aecInputDelay = 2; // in 10ms blocks
	string                      aecDumpFileDir;

	obs_source_t                *source;
	string                      device_id;
	string                      device_name;
	bool                        isInputDevice;
	bool                        useDeviceTiming = false;
	bool                        isDefaultDevice = false;

	bool                        reconnecting = false;
	bool                        previouslyFailed = false;
	WinHandle                   reconnectThread;

	bool                        active = false;
	WinHandle                   captureThread;

	WinHandle                   stopSignal;
	WinHandle                   receiveSignal;

	speaker_layout              speakers;
	audio_format                format;
	uint32_t                    sampleRate;

	static DWORD WINAPI ReconnectThread(LPVOID param);
	static DWORD WINAPI CaptureThread(LPVOID param);

	bool ProcessCaptureData(bool& dmoActive, deque<ComPtr<IMediaBuffer>>& inputQueue,
		FILE* pcmDumpInput, FILE* pcmDumpLoopback, FILE* pcmDumpOutput);

	inline void Start();
	inline void Stop();
	void Reconnect();

	bool InitDevice(IMMDeviceEnumerator *enumerator);
	void InitName();
	void InitClient();
	void InitRender();
	void InitFormat(WAVEFORMATEX *wfex);
	void InitCapture();
	void Initialize();

	bool TryInitialize();

	void UpdateSettings(obs_data_t *settings);

public:
	WASAPISource(obs_data_t *settings, obs_source_t *source_, bool input);
	inline ~WASAPISource();

	void Update(obs_data_t *settings);
};

WASAPISource::WASAPISource(obs_data_t *settings, obs_source_t *source_,
		bool input)
	: source          (source_),
	  isInputDevice   (input)
{
	UpdateSettings(settings);

	stopSignal = CreateEvent(nullptr, true, false, nullptr);
	if (!stopSignal.Valid())
		throw "Could not create stop signal";

	receiveSignal = CreateEvent(nullptr, false, false, nullptr);
	if (!receiveSignal.Valid())
		throw "Could not create receive signal";

	Start();
}

inline void WASAPISource::Start()
{
	if (!TryInitialize()) {
		blog(LOG_INFO, "[WASAPISource::WASAPISource] "
		               "Device '%s' not found.  Waiting for device",
		               device_id.c_str());
		Reconnect();
	}
}

inline void WASAPISource::Stop()
{
	SetEvent(stopSignal);

	if (active) {
		blog(LOG_INFO, "WASAPI: Device '%s' Terminated",
				device_name.c_str());
		WaitForSingleObject(captureThread, INFINITE);
	}

	if (reconnecting)
		WaitForSingleObject(reconnectThread, INFINITE);

	ResetEvent(stopSignal);
}

inline WASAPISource::~WASAPISource()
{
	Stop();
}

void WASAPISource::UpdateSettings(obs_data_t *settings)
{
	device_id       = obs_data_get_string(settings, OPT_DEVICE_ID);
	useDeviceTiming = obs_data_get_bool(settings, OPT_USE_DEVICE_TIMING);
	isDefaultDevice = _strcmpi(device_id.c_str(), "default") == 0;
	disableAEC = obs_data_get_bool(settings, OPT_DISABLE_AEC);
	inFormatMode = (int)obs_data_get_int(settings, OPT_IN_FMT_MODE);
	aecInputDelay = (int)obs_data_get_int(settings, OPT_AEC_IN_DELAY);
	aecDumpFileDir = obs_data_get_string(settings, OPT_AEC_DUMP_DIR);

	blog(LOG_INFO, "disable_aec: %d, input delay: %d, dump dir: %s, fmt_mode: %d",
		disableAEC, aecInputDelay, aecDumpFileDir.c_str(), inFormatMode);
}

void WASAPISource::Update(obs_data_t *settings)
{
	string newDevice = obs_data_get_string(settings, OPT_DEVICE_ID);
	bool newDisableAEC = obs_data_get_bool(settings, OPT_DISABLE_AEC);
	int newInFormatMode = (int)obs_data_get_int(settings, OPT_IN_FMT_MODE);
	int newInputDelay = (int)obs_data_get_int(settings, OPT_AEC_IN_DELAY);
	string newDumpFileDir = obs_data_get_string(settings, OPT_AEC_DUMP_DIR);
	bool restart = (
		newDevice.compare(device_id) != 0 ||
		newDisableAEC != disableAEC ||
		newInFormatMode != inFormatMode ||
		newInputDelay != aecInputDelay ||
		newDumpFileDir.compare(aecDumpFileDir) != 0);

	if (restart)
		Stop();

	UpdateSettings(settings);

	if (restart)
		Start();
}

static inline HRESULT GetMixFormat(
	ComPtr<IAudioClient> client, WAVEFORMATEX **ppFormat, const char* clientName)
{
	HRESULT res = client->GetMixFormat(ppFormat);
	if (FAILED(res))
		return res;
	blog(LOG_INFO, "MixFormat (%s) ch: %d, bits: %d, sampleRate: %d, formatTag: %d",
		clientName, (*ppFormat)->nChannels, (*ppFormat)->wBitsPerSample,
		(*ppFormat)->nSamplesPerSec, (*ppFormat)->wFormatTag);

	WAVEFORMATEX* closestFormat = nullptr;
	res = client->IsFormatSupported(AUDCLNT_SHAREMODE_SHARED, *ppFormat, &closestFormat);
	if (res != S_FALSE) // succeeded, or failed with no alternative
		return res;

	CoTaskMemFree(*ppFormat);
	*ppFormat = closestFormat;
	blog(LOG_INFO, "ClosestFormat (%s) ch: %d, bits: %d, sampleRate: %d, formatTag: %d",
		clientName, closestFormat->nChannels, closestFormat->wBitsPerSample,
		closestFormat->nSamplesPerSec, closestFormat->wFormatTag);
	return S_OK;
}

bool WASAPISource::InitDevice(IMMDeviceEnumerator *enumerator)
{
	HRESULT res;

	if (isDefaultDevice) {
		res = enumerator->GetDefaultAudioEndpoint(
				isInputDevice ? eCapture        : eRender,
				isInputDevice ? eCommunications : eConsole,
				device.Assign());
	} else {
		wchar_t *w_id;
		os_utf8_to_wcs_ptr(device_id.c_str(), device_id.size(), &w_id);

		res = enumerator->GetDevice(w_id, device.Assign());

		bfree(w_id);
	}

	if (isInputDevice && SUCCEEDED(res) && !disableAEC) {
		enumerator->GetDefaultAudioEndpoint(
			eRender, eConsole, deviceRender.Assign());
	}

	return SUCCEEDED(res);
}

#define BUFFER_TIME_100NS (5*10000000)

void WASAPISource::InitClient()
{
	HRESULT res;
	DWORD   flags = AUDCLNT_STREAMFLAGS_EVENTCALLBACK;
	int     pass = 0;

	res = device->Activate(__uuidof(IAudioClient), CLSCTX_ALL,
			nullptr, (void**)client.Assign());
	if (FAILED(res))
		throw HRError("Failed to activate client context", res);

	res = GetMixFormat(client, &wfexClient, "client");
	if (FAILED(res))
		throw HRError("Failed to get mix format", res);

	if (!isInputDevice)
		flags |= AUDCLNT_STREAMFLAGS_LOOPBACK;

	while (true) {
		if (pass >= inFormatMode) {
			InitFormat(wfexClient);

			res = client->Initialize(
				AUDCLNT_SHAREMODE_SHARED, flags,
				BUFFER_TIME_100NS, 0, wfexClient, nullptr);

			if (SUCCEEDED(res)) {
				break;
			}
		}

		pass ++;
		switch (pass) {
			case 1:
				wfexClient->nChannels = 1;
				wfexClient->wFormatTag = WAVE_FORMAT_PCM;
				wfexClient->wBitsPerSample = 16;
				wfexClient->cbSize = 0;
				flags |= AUDCLNT_STREAMFLAGS_AUTOCONVERTPCM |
				         AUDCLNT_STREAMFLAGS_SRC_DEFAULT_QUALITY;
				break;
			case 2:
				if (isInputDevice && !disableAEC) {
					wfexClient->nSamplesPerSec = 22050;
				}
				else {
					wfexClient->nSamplesPerSec = 44100;
				}
				break;
			case 3:
				throw HRError("Failed to initialize audio client", res);
		}
		wfexClient->nBlockAlign = wfexClient->nChannels * (wfexClient->wBitsPerSample / 8);
		wfexClient->nAvgBytesPerSec = wfexClient->nSamplesPerSec * wfexClient->nBlockAlign;

		blog(LOG_INFO, "Re-initialize audio client on error (%lX), pass %d", res, pass);
		client.Clear();
		res = device->Activate(__uuidof(IAudioClient), CLSCTX_ALL,
			nullptr, (void**)client.Assign());
		if (FAILED(res))
			throw HRError("Failed to activate client context", res);
	}
}

void WASAPISource::InitRender()
{
	HRESULT res;
	DWORD   flags = 0;
	LPBYTE  buffer;
	UINT32  frames;
	int     pass = 0;

	if (isInputDevice) {
		if (!deviceRender.Get())
			return;
		res = deviceRender->Activate(__uuidof(IAudioClient), CLSCTX_ALL,
			nullptr, (void**)clientRender.Assign());
	}
	else {
		res = device->Activate(__uuidof(IAudioClient), CLSCTX_ALL,
			nullptr, (void**)clientRender.Assign());
	}

	if (FAILED(res))
		throw HRError("Failed to activate client context", res);

	res = GetMixFormat(clientRender, &wfexClientRender, "clientRender");
	if (FAILED(res))
		throw HRError("Failed to get mix format", res);

	if (isInputDevice)
		flags |= AUDCLNT_STREAMFLAGS_LOOPBACK;

	while (true) {
		if (pass >= inFormatMode) {
			res = clientRender->Initialize(
				AUDCLNT_SHAREMODE_SHARED, flags,
				BUFFER_TIME_100NS, 0, wfexClientRender, nullptr);

			if (SUCCEEDED(res)) {
				break;
			}
		}

		pass ++;
		switch (pass) {
			case 1:
				wfexClientRender->nChannels = 1;
				wfexClientRender->wFormatTag = WAVE_FORMAT_PCM;
				wfexClientRender->wBitsPerSample = 16;
				wfexClientRender->cbSize = 0;
				flags |= AUDCLNT_STREAMFLAGS_AUTOCONVERTPCM |
				         AUDCLNT_STREAMFLAGS_SRC_DEFAULT_QUALITY;
				break;
			case 2:
				if (isInputDevice) {
					wfexClientRender->nSamplesPerSec = 22050;
				}
				else {
					wfexClientRender->nSamplesPerSec = 44100;
				}
				break;
			case 3:
				throw HRError("Failed to initialize audio client", res);
		}
		wfexClientRender->nBlockAlign = wfexClientRender->nChannels * (wfexClientRender->wBitsPerSample / 8);
		wfexClientRender->nAvgBytesPerSec = wfexClientRender->nSamplesPerSec * wfexClientRender->nBlockAlign;

		blog(LOG_INFO, "Re-initialize audio render client on error (%lX), pass %d", res, pass);
		clientRender.Clear();
		if (isInputDevice) {
			res = deviceRender->Activate(__uuidof(IAudioClient), CLSCTX_ALL,
				nullptr, (void**)clientRender.Assign());
		}
		else {
			res = device->Activate(__uuidof(IAudioClient), CLSCTX_ALL,
				nullptr, (void**)clientRender.Assign());
		}
		if (FAILED(res))
			throw HRError("Failed to activate client context", res);
	}

	if (isInputDevice)
		return;

	/* Silent loopback fix. Prevents audio stream from stopping and */
	/* messing up timestamps and other weird glitches during silence */
	/* by playing a silent sample all over again. */

	res = clientRender->GetBufferSize(&frames);
	if (FAILED(res))
		throw HRError("Failed to get buffer size", res);

	res = clientRender->GetService(__uuidof(IAudioRenderClient),
		(void**)render.Assign());
	if (FAILED(res))
		throw HRError("Failed to get render client", res);

	res = render->GetBuffer(frames, &buffer);
	if (FAILED(res))
		throw HRError("Failed to get buffer", res);

	memset(buffer, 0, frames*wfexClientRender->nBlockAlign);

	render->ReleaseBuffer(frames, 0);
}

static speaker_layout ConvertSpeakerLayout(DWORD layout, WORD channels)
{
	switch (layout) {
	case KSAUDIO_SPEAKER_2POINT1:          return SPEAKERS_2POINT1;
	case KSAUDIO_SPEAKER_SURROUND:         return SPEAKERS_4POINT0;
	case OBS_KSAUDIO_SPEAKER_4POINT1:      return SPEAKERS_4POINT1;
	case KSAUDIO_SPEAKER_5POINT1_SURROUND: return SPEAKERS_5POINT1;
	case KSAUDIO_SPEAKER_7POINT1_SURROUND: return SPEAKERS_7POINT1;
	}

	return (speaker_layout)channels;
}

void WASAPISource::InitFormat(WAVEFORMATEX *wfex)
{
	DWORD layout = 0;

	if (wfex->wFormatTag == WAVE_FORMAT_EXTENSIBLE) {
		WAVEFORMATEXTENSIBLE *ext = (WAVEFORMATEXTENSIBLE*)wfex;
		layout = ext->dwChannelMask;
	}

	/* WASAPI is always float */
	sampleRate = wfex->nSamplesPerSec;
	format     = AUDIO_FORMAT_FLOAT;
	speakers   = ConvertSpeakerLayout(layout, wfex->nChannels);

	if (wfex->wFormatTag == WAVE_FORMAT_PCM) {
		assert(wfex->wBitsPerSample == 16);
		format = AUDIO_FORMAT_16BIT;
	}

	blog(LOG_INFO, "##### Device Type: %s, channels: %d, bitspersample: %d, samplerate: %d",
		isInputDevice ? "input": "output",
		wfex->nChannels, wfex->wBitsPerSample, wfex->nSamplesPerSec);
}

void WASAPISource::InitCapture()
{
	HRESULT res = client->GetService(__uuidof(IAudioCaptureClient),
			(void**)capture.Assign());
	if (FAILED(res))
		throw HRError("Failed to create capture context", res);

	res = client->SetEventHandle(receiveSignal);
	if (FAILED(res))
		throw HRError("Failed to set event handle", res);

	blog(LOG_INFO, "InitCapture: %d", (int)isInputDevice);

	if (isInputDevice && !disableAEC && clientRender.Get()) {
		try {
			res = clientRender->GetService(__uuidof(IAudioCaptureClient),
				(void**)captureRender.Assign());
			if (FAILED(res))
				throw HRError("Failed to create render capture context", res);

			res = CoCreateInstance(__uuidof(CWMAudioAEC),
				nullptr, CLSCTX_INPROC_SERVER,
				__uuidof(IMediaObject),
				(void**)captureDMO.Assign());

			if (FAILED(res))
				throw HRError("Failed to create capture DMO", res);

			ComPtr <IPropertyStore> dmoProp;

			res = captureDMO->QueryInterface(__uuidof(IPropertyStore),
				(void**)dmoProp.Assign());

			if (FAILED(res))
				throw HRError("Failed to get dmo prop", res);

			PROPVARIANT pv;
			PropVariantInit(&pv);
			pv.vt = VT_BOOL;
			pv.lVal = VARIANT_FALSE;

			res = dmoProp->SetValue(MFPKEY_WMAAECMA_DMO_SOURCE_MODE, pv);
			if (FAILED(res))
				throw HRError("Failed to enable filter mode", res);

			PropVariantInit(&pv);
			pv.vt = VT_I4;
			pv.lVal = (LONG)0; // AEC only

			res = dmoProp->SetValue(MFPKEY_WMAAECMA_SYSTEM_MODE, pv);
			if (FAILED(res))
				throw HRError("Failed to set dmo system mode", res);

#if 0 // experimental: changing AEC echo length setting
			PropVariantInit(&pv);
			pv.vt = VT_BOOL;
			pv.lVal = VARIANT_TRUE;

			res = dmoProp->SetValue(MFPKEY_WMAAECMA_FEATURE_MODE, pv);
			if (FAILED(res))
				throw HRError("Failed to enable feature mode", res);

			PropVariantInit(&pv);
			pv.vt = VT_I4;
			pv.lVal = (LONG)1024; // Max

			res = dmoProp->SetValue(MFPKEY_WMAAECMA_FEATR_ECHO_LENGTH, pv);
			if (FAILED(res))
				throw HRError("Failed to set echo length", res);
#endif

			DMO_MEDIA_TYPE mt;

			CoTaskMemPtr<WAVEFORMATEX> wfex;
			*(&wfex) = (WAVEFORMATEX*)CoTaskMemAlloc(sizeof(WAVEFORMATEX));
			assert(wfex); // it should never run out of memory

			mt.majortype = MEDIATYPE_Audio;
			mt.subtype = MEDIASUBTYPE_PCM;
			mt.lSampleSize = 0;
			mt.bFixedSizeSamples = TRUE;
			mt.bTemporalCompression = FALSE;
			mt.formattype = FORMAT_WaveFormatEx;
			mt.cbFormat = sizeof(WAVEFORMATEX);
			mt.pbFormat = (BYTE*)(WAVEFORMATEX*)wfex;

			wfex->wFormatTag = WAVE_FORMAT_PCM;
			wfex->nSamplesPerSec = wfexClient->nSamplesPerSec;
			wfex->nAvgBytesPerSec = wfexClient->nSamplesPerSec * 2;
			wfex->nChannels = 1;
			wfex->nBlockAlign = 2;
			wfex->wBitsPerSample = 16;
			wfex->cbSize = 0;

			res = captureDMO->SetInputType(0, &mt, 0);
			if (FAILED(res))
				throw HRError("Failed to set input type 0", res);

			wfex->nSamplesPerSec = wfexClientRender->nSamplesPerSec;
			wfex->nAvgBytesPerSec = wfexClientRender->nSamplesPerSec * 2;

			res = captureDMO->SetInputType(1, &mt, 0);
			if (FAILED(res))
				throw HRError("Failed to set input type 1", res);

			wfex->nSamplesPerSec = 22050;
			wfex->nAvgBytesPerSec = 22050 * 2;

			res = captureDMO->SetOutputType(0, &mt, 0);
			if (FAILED(res))
				throw HRError("Failed to set dmo output type", res);

			res = captureDMO->AllocateStreamingResources();
			if (FAILED(res))
				throw HRError("Failed to allocate dmo streaming resource", res);

			res = CMediaBuffer::Create(2 * 2 * 22050, 0, captureDMOBuffer.Assign());
			if (FAILED(res))
				throw HRError("Failed to allocate dmo output buffer", res);

			blog(LOG_INFO, "DMO init success");
		}
		catch (...) {
			captureDMO = NULL;
			captureDMOBuffer = NULL;
			blog(LOG_WARNING, "WASAPI: Failed to config AEC DMO");
		}
	}
	else {
		blog(LOG_INFO, "DMO NOT initialized, AEC disabled");
	}

	captureThread = CreateThread(nullptr, 0,
			WASAPISource::CaptureThread, this,
			0, nullptr);
	if (!captureThread.Valid())
		throw "Failed to create capture thread";

	client->Start();

	if (captureDMO.Get()) {
		clientRender->Start();
	}

	active = true;

	blog(LOG_INFO, "WASAPI: Device '%s' initialized", device_name.c_str());
}

void WASAPISource::Initialize()
{
	ComPtr<IMMDeviceEnumerator> enumerator;
	HRESULT res;

	device.Clear();
	deviceRender.Clear();
	client.Clear();
	clientRender.Clear();
	capture.Clear();
	captureRender.Clear();
	render.Clear();
	captureDMO.Clear();
	captureDMOBuffer.Clear();
	*(&wfexClient) = nullptr;
	*(&wfexClientRender) = nullptr;

	res = CoCreateInstance(__uuidof(MMDeviceEnumerator),
			nullptr, CLSCTX_ALL,
			__uuidof(IMMDeviceEnumerator),
			(void**)enumerator.Assign());
	if (FAILED(res))
		throw HRError("Failed to create enumerator", res);

	if (!InitDevice(enumerator))
		return;

	device_name = GetDeviceName(device);

	InitClient();

	try {
		InitRender();
	}
	catch (HRError error) {
		if (isInputDevice) {
			clientRender.Clear();
			blog(LOG_WARNING, "Ignored loopback init error - %s: %lX", error.str, error.hr);
		}
		else {
			throw error;
		}
	}

	InitCapture();
}

bool WASAPISource::TryInitialize()
{
	try {
		Initialize();

	} catch (HRError error) {
		if (previouslyFailed)
			return active;

		blog(LOG_WARNING, "[WASAPISource::TryInitialize]:[%s] %s: %lX",
				device_name.empty() ?
					device_id.c_str() : device_name.c_str(),
				error.str, error.hr);

	} catch (const char *error) {
		if (previouslyFailed)
			return active;

		blog(LOG_WARNING, "[WASAPISource::TryInitialize]:[%s] %s",
				device_name.empty() ?
					device_id.c_str() : device_name.c_str(),
				error);
	}

	previouslyFailed = !active;
	return active;
}

void WASAPISource::Reconnect()
{
	reconnecting = true;
	reconnectThread = CreateThread(nullptr, 0,
			WASAPISource::ReconnectThread, this,
			0, nullptr);

	if (!reconnectThread.Valid())
		blog(LOG_WARNING, "[WASAPISource::Reconnect] "
		                "Failed to initialize reconnect thread: %lu",
		                 GetLastError());
}

static inline bool WaitForSignal(HANDLE handle, DWORD time)
{
	return WaitForSingleObject(handle, time) != WAIT_TIMEOUT;
}

#define RECONNECT_INTERVAL 3000

DWORD WINAPI WASAPISource::ReconnectThread(LPVOID param)
{
	WASAPISource *source = (WASAPISource*)param;

	os_set_thread_name("win-wasapi: reconnect thread");

	CoInitializeEx(0, COINIT_MULTITHREADED);

	obs_monitoring_type type = obs_source_get_monitoring_type(source->source);
	obs_source_set_monitoring_type(source->source, OBS_MONITORING_TYPE_NONE);

	while (!WaitForSignal(source->stopSignal, RECONNECT_INTERVAL)) {
		if (source->TryInitialize())
			break;
	}

	obs_source_set_monitoring_type(source->source, type);

	source->reconnectThread = nullptr;
	source->reconnecting = false;
	return 0;
}

bool WASAPISource::ProcessCaptureData(bool& dmoActive, deque<ComPtr<IMediaBuffer>>& inputQueue,
	FILE* pcmDumpInput, FILE* pcmDumpLoopback, FILE* pcmDumpOutput)
{
	HRESULT res;
	LPBYTE  buffer;
	UINT32  frames;
	DWORD   flags;
	UINT64  pos, ts;
	UINT    captureSize = 0;

	while (true) {
		res = capture->GetNextPacketSize(&captureSize);

		if (FAILED(res)) {
			if (res != AUDCLNT_E_DEVICE_INVALIDATED)
				blog(LOG_WARNING,
						"[WASAPISource::GetCaptureData]"
						" capture->GetNextPacketSize"
						" failed: %lX", res);
			return false;
		}

		if (!captureSize)
			break;

		res = capture->GetBuffer(&buffer, &frames, &flags, &pos, &ts);
		if (FAILED(res)) {
			if (res != AUDCLNT_E_DEVICE_INVALIDATED)
				blog(LOG_WARNING,
						"[WASAPISource::GetCaptureData]"
						" capture->GetBuffer"
						" failed: %lX", res);
			return false;
		}

		if (!frames) {
			capture->ReleaseBuffer(frames);
			break;
		}

		obs_source_audio data = {};

		if (captureDMO.Get()) {
			ComPtr<IMediaBuffer> inMediaBuffer, outMediaBuffer;
			BYTE* pData;

			if (SUCCEEDED(CMediaBuffer::Create(frames * 2, ts, inMediaBuffer.Assign()))) {
				inMediaBuffer->SetLength(frames * 2);
				inMediaBuffer->GetBufferAndLength(&pData, NULL);
				if (wfexClient->wFormatTag == WAVE_FORMAT_PCM) {
					memcpy(pData, buffer, frames * 2);
				}
				else {
					int nChannels = wfexClient->nChannels;
					for (UINT32 i = 0; i < frames; i++) {
						float f = ((float*)buffer)[i * nChannels];
						((short*)pData)[i] = (short)(((f > 1.0f)? 1.0f : (f < -1.0f)? -1.0f : f) * 0x7fff);
					}
				}
				inputQueue.push_back(inMediaBuffer);
			}
			else {
				assert(false); // it should never run out of memory
			}

			LPBYTE outbuffer;
			UINT32 outframes = 0;
			UINT64 outPos, outTs;

			captureSize = 0;
			if (SUCCEEDED(captureRender->GetNextPacketSize(&captureSize)) && captureSize &&
				SUCCEEDED(captureRender->GetBuffer(&outbuffer, &outframes, &flags, &outPos, &outTs))) {

				if (outframes && SUCCEEDED(CMediaBuffer::Create(outframes * 2, 0, outMediaBuffer.Assign()))) {
					outMediaBuffer->SetLength(outframes * 2);
					outMediaBuffer->GetBufferAndLength(&pData, NULL);
					if (wfexClientRender->wFormatTag == WAVE_FORMAT_PCM) {
						memcpy(pData, outbuffer, outframes * 2);
					}
					else {
						int nChannels = wfexClientRender->nChannels;
						for (UINT32 i = 0; i < outframes; i++) {
							float f = ((float*)outbuffer)[i * nChannels];
							((short*)pData)[i] = (short)(((f > 1.0f)? 1.0f : (f < -1.0f)? -1.0f : f) * 0x7fff);
						}
					}
				}
				else {
					assert(!outframes); // it should never run out of memory
				}

				captureRender->ReleaseBuffer(outframes);
			}

			if (inputQueue.size() > aecInputDelay) {
				bool outputDone = false;

				inMediaBuffer = inputQueue.front();
				inputQueue.pop_front();

				if (outMediaBuffer) {
					if (pcmDumpInput) { ((CMediaBuffer*)inMediaBuffer.Get())->Dump(pcmDumpInput); }
					if (pcmDumpLoopback) { ((CMediaBuffer*)outMediaBuffer.Get())->Dump(pcmDumpLoopback); }

					if (!dmoActive) {
						captureDMO->Flush();
						blog(LOG_INFO, "DMO Flush");
					}

					do {
						res = captureDMO->ProcessInput(0, inMediaBuffer,
							DMO_INPUT_DATA_BUFFERF_SYNCPOINT | DMO_INPUT_DATA_BUFFERF_TIME,
							((CMediaBuffer*)inMediaBuffer.Get())->GetTimestamp(), 0);
						if (FAILED(res)) {
							blog(LOG_ERROR, "Failed to process dmo input 0 %x", res);
							break;
						}

						res = captureDMO->ProcessInput(1, outMediaBuffer,
							DMO_INPUT_DATA_BUFFERF_SYNCPOINT | DMO_INPUT_DATA_BUFFERF_TIME,
							outTs, 0);
						if (FAILED(res)) {
							blog(LOG_ERROR, "Failed to process dmo input 1 %x", res);
							break;
						}

						DMO_OUTPUT_DATA_BUFFER dmoOut;
						DWORD dmoStatus;
						dmoOut.pBuffer = captureDMOBuffer;
						captureDMOBuffer->SetLength(0);

						res = captureDMO->ProcessOutput(0, 1, &dmoOut, &dmoStatus);
						if (FAILED(res)) {
							blog(LOG_ERROR, "Failed to process dmo output %x", res);
							break;
						}

						dmoActive = outputDone = true;

						DWORD len;
						captureDMOBuffer->GetBufferAndLength(&pData, &len);

						if (len > 1) {
							data.data[0] = (const uint8_t*)pData;
							data.frames = len / 2;
							data.speakers = SPEAKERS_MONO;
							data.samples_per_sec = 22050;
							data.format = AUDIO_FORMAT_16BIT;
							data.timestamp = os_gettime_ns() -
								(uint64_t)data.frames * 1000000000ULL /
								(uint64_t)data.samples_per_sec;

							obs_source_output_audio(source, &data);

							if (pcmDumpOutput) { fwrite(pData, 1, len, pcmDumpOutput); }
						}
					}
					while (false);
				}

				if (!outputDone) {
					dmoActive = false;

					DWORD len;
					inMediaBuffer->GetBufferAndLength(&pData, &len);

					if (len > 1) {
						data.data[0] = (const uint8_t*)pData;
						data.frames = len / 2;
						data.speakers = SPEAKERS_MONO;
						data.samples_per_sec = wfexClient->nSamplesPerSec;
						data.format = AUDIO_FORMAT_16BIT;
						data.timestamp = os_gettime_ns() -
							(uint64_t)data.frames * 1000000000ULL /
							(uint64_t)data.samples_per_sec;

						obs_source_output_audio(source, &data);
					}
				}
			}
		}
		else {
			data.data[0] = (const uint8_t*)buffer;
			data.frames = (uint32_t)frames;
			data.speakers = speakers;
			data.samples_per_sec = sampleRate;
			data.format = format;
			data.timestamp = useDeviceTiming ?
				ts * 100 : os_gettime_ns();

			if (!useDeviceTiming)
				data.timestamp -= (uint64_t)frames * 1000000000ULL /
				(uint64_t)sampleRate;

			obs_source_output_audio(source, &data);
		}

		capture->ReleaseBuffer(frames);
	}

	return true;
}

static inline bool WaitForCaptureSignal(DWORD numSignals, const HANDLE *signals,
		DWORD duration)
{
	DWORD ret;
	ret = WaitForMultipleObjects(numSignals, signals, false, duration);

	return ret == WAIT_OBJECT_0 || ret == WAIT_TIMEOUT;
}

DWORD WINAPI WASAPISource::CaptureThread(LPVOID param)
{
	WASAPISource *source   = (WASAPISource*)param;
	bool         reconnect = false;

	/* Output devices don't signal, so just make it check every 10 ms */
	DWORD        dur       = source->isInputDevice ? RECONNECT_INTERVAL : 10;

	HANDLE sigs[2] = {
		source->receiveSignal,
		source->stopSignal
	};

	os_set_thread_name("win-wasapi: capture thread");

	FILE* pcmDumpInput = NULL;
	FILE* pcmDumpLoopback = NULL;
	FILE* pcmDumpOutput = NULL;
	bool dmoActive = false;

	deque<ComPtr<IMediaBuffer>> inputQueue;

	if (source->captureDMO.Get() && !source->aecDumpFileDir.empty()) {
		time_t tm = time(NULL);
		pcmDumpInput = fopen((source->aecDumpFileDir+
			"/aec_in0-"+to_string(tm)+".pcm").c_str(), "wb");
		pcmDumpLoopback = fopen((source->aecDumpFileDir+
			"/aec_in1-"+to_string(tm)+".pcm").c_str(), "wb");
		pcmDumpOutput = fopen((source->aecDumpFileDir+
			"/aec_out-"+to_string(tm)+".pcm").c_str(), "wb");
	}

	while (WaitForCaptureSignal(2, sigs, dur)) {
		if (!source->ProcessCaptureData(dmoActive, inputQueue,
			pcmDumpInput, pcmDumpLoopback, pcmDumpOutput)) {
			reconnect = true;
			break;
		}
	}

	if (pcmDumpInput)    { fclose(pcmDumpInput); }
	if (pcmDumpLoopback) { fclose(pcmDumpLoopback); }
	if (pcmDumpOutput)   { fclose(pcmDumpOutput); }

	source->client->Stop();

	source->captureThread = nullptr;
	source->active        = false;

	if (reconnect) {
		blog(LOG_INFO, "Device '%s' invalidated.  Retrying",
				source->device_name.c_str());
		source->Reconnect();
	}

	return 0;
}

/* ------------------------------------------------------------------------- */

static const char *GetWASAPIInputName(void*)
{
	return obs_module_text("AudioInput");
}

static const char *GetWASAPIOutputName(void*)
{
	return obs_module_text("AudioOutput");
}

static void GetWASAPIDefaultsInput(obs_data_t *settings)
{
	obs_data_set_default_string(settings, OPT_DEVICE_ID, "default");
	obs_data_set_default_bool(settings, OPT_USE_DEVICE_TIMING, false);
	obs_data_set_default_bool(settings, OPT_DISABLE_AEC, false);
	obs_data_set_default_int(settings, OPT_IN_FMT_MODE, 0);
	obs_data_set_default_int(settings, OPT_AEC_IN_DELAY, 2);
	obs_data_set_default_string(settings, OPT_AEC_DUMP_DIR, "");
}

static void GetWASAPIDefaultsOutput(obs_data_t *settings)
{
	obs_data_set_default_string(settings, OPT_DEVICE_ID, "default");
	obs_data_set_default_bool(settings, OPT_USE_DEVICE_TIMING, true);
	obs_data_set_default_bool(settings, OPT_DISABLE_AEC, false);
	obs_data_set_default_int(settings, OPT_IN_FMT_MODE, 0);
	obs_data_set_default_int(settings, OPT_AEC_IN_DELAY, 2);
	obs_data_set_default_string(settings, OPT_AEC_DUMP_DIR, "");
}

static void *CreateWASAPISource(obs_data_t *settings, obs_source_t *source,
		bool input)
{
	try {
		return new WASAPISource(settings, source, input);
	} catch (const char *error) {
		blog(LOG_ERROR, "[CreateWASAPISource] %s", error);
	}

	return nullptr;
}

static void *CreateWASAPIInput(obs_data_t *settings, obs_source_t *source)
{
	return CreateWASAPISource(settings, source, true);
}

static void *CreateWASAPIOutput(obs_data_t *settings, obs_source_t *source)
{
	return CreateWASAPISource(settings, source, false);
}

static void DestroyWASAPISource(void *obj)
{
	delete static_cast<WASAPISource*>(obj);
}

static void UpdateWASAPISource(void *obj, obs_data_t *settings)
{
	static_cast<WASAPISource*>(obj)->Update(settings);
}

static obs_properties_t *GetWASAPIProperties(bool input)
{
	obs_properties_t *props = obs_properties_create();
	vector<AudioDeviceInfo> devices;

	obs_property_t *device_prop = obs_properties_add_list(props,
			OPT_DEVICE_ID, obs_module_text("Device"),
			OBS_COMBO_TYPE_LIST, OBS_COMBO_FORMAT_STRING);

	GetWASAPIAudioDevices(devices, input);

	if (devices.size())
		obs_property_list_add_string(device_prop,
				obs_module_text("Default"), "default");

	for (size_t i = 0; i < devices.size(); i++) {
		AudioDeviceInfo &device = devices[i];
		obs_property_list_add_string(device_prop,
				device.name.c_str(), device.id.c_str());
	}

	obs_properties_add_bool(props, OPT_USE_DEVICE_TIMING,
			obs_module_text("UseDeviceTiming"));

	obs_properties_add_bool(props,
		OPT_DISABLE_AEC, "Disable Echo Cancellation");
	obs_properties_add_int(props,
		OPT_IN_FMT_MODE, "Audio Input Format Mode", 0, 3, 1);
	obs_properties_add_int(props,
		OPT_AEC_IN_DELAY, "AEC Input Delay", 0, 9, 1);
	obs_properties_add_path(props,
		OPT_AEC_DUMP_DIR, "AEC Dump File Dir", OBS_PATH_DIRECTORY, NULL, NULL);

	return props;
}

static obs_properties_t *GetWASAPIPropertiesInput(void *)
{
	return GetWASAPIProperties(true);
}

static obs_properties_t *GetWASAPIPropertiesOutput(void *)
{
	return GetWASAPIProperties(false);
}

void RegisterWASAPIInput()
{
	obs_source_info info = {};
	info.id              = "wasapi_input_capture";
	info.type            = OBS_SOURCE_TYPE_INPUT;
	info.output_flags    = OBS_SOURCE_AUDIO |
	                       OBS_SOURCE_DO_NOT_DUPLICATE;
	info.get_name        = GetWASAPIInputName;
	info.create          = CreateWASAPIInput;
	info.destroy         = DestroyWASAPISource;
	info.update          = UpdateWASAPISource;
	info.get_defaults    = GetWASAPIDefaultsInput;
	info.get_properties  = GetWASAPIPropertiesInput;
	obs_register_source(&info);
}

void RegisterWASAPIOutput()
{
	obs_source_info info = {};
	info.id              = "wasapi_output_capture";
	info.type            = OBS_SOURCE_TYPE_INPUT;
	info.output_flags    = OBS_SOURCE_AUDIO |
	                       OBS_SOURCE_DO_NOT_DUPLICATE |
	                       OBS_SOURCE_DO_NOT_SELF_MONITOR;
	info.get_name        = GetWASAPIOutputName;
	info.create          = CreateWASAPIOutput;
	info.destroy         = DestroyWASAPISource;
	info.update          = UpdateWASAPISource;
	info.get_defaults    = GetWASAPIDefaultsOutput;
	info.get_properties  = GetWASAPIPropertiesOutput;
	obs_register_source(&info);
}
