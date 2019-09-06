#include "enum-wasapi.hpp"

#include <obs-module.h>
#include <obs.h>
#include <util/platform.h>
#include <util/windows/HRError.hpp>
#include <util/windows/ComPtr.hpp>
#include <util/windows/WinHandle.hpp>
#include <util/windows/CoTaskMemPtr.hpp>
#include <util/threading.h>

using namespace std;

#define OPT_DEVICE_ID         "device_id"
#define OPT_USE_DEVICE_TIMING "use_device_timing"

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

	CMediaBuffer(DWORD cbMaxLength, HRESULT& hr) :
		m_nRefCount(1),
		m_cbMaxLength(cbMaxLength),
		m_cbLength(0),
		m_pbData(NULL)
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
	static HRESULT Create(long cbMaxLen, IMediaBuffer **ppBuffer)
	{
		HRESULT hr = S_OK;
		CMediaBuffer *pBuffer = NULL;

		if (ppBuffer == NULL)
		{
			return E_POINTER;
		}

		pBuffer = new CMediaBuffer(cbMaxLen, hr);

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

	bool                        disableAEC = false;
	int                         nChannels;
	int                         nChannelsRender;
	int                         nSampleRate;
	int                         nSampleRateRender;

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

	bool ProcessCaptureData(bool& dmoActive, FILE* pcmDumpInput, FILE* pcmDumpLoopback, FILE* pcmDumpOutput);

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
	disableAEC = obs_data_get_bool(settings, "disable_echo_cancellation");
}

void WASAPISource::Update(obs_data_t *settings)
{
	string newDevice = obs_data_get_string(settings, OPT_DEVICE_ID);
	bool newDisableAEC = obs_data_get_bool(settings, "disable_echo_cancellation");
	bool restart = newDevice.compare(device_id) != 0;

	blog(LOG_INFO, "disable_echo: %d", newDisableAEC);

	if (newDisableAEC != disableAEC) {
		restart = true;
	}

	if (restart)
		Stop();

	UpdateSettings(settings);

	if (restart)
		Start();
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
	CoTaskMemPtr<WAVEFORMATEX> wfex;
	HRESULT                    res;
	DWORD                      flags = AUDCLNT_STREAMFLAGS_EVENTCALLBACK;

	res = device->Activate(__uuidof(IAudioClient), CLSCTX_ALL,
			nullptr, (void**)client.Assign());
	if (FAILED(res))
		throw HRError("Failed to activate client context", res);

	res = client->GetMixFormat(&wfex);
	if (FAILED(res))
		throw HRError("Failed to get mix format", res);

	InitFormat(wfex);

	if (!isInputDevice)
		flags |= AUDCLNT_STREAMFLAGS_LOOPBACK;

	res = client->Initialize(
			AUDCLNT_SHAREMODE_SHARED, flags,
			BUFFER_TIME_100NS, 0, wfex, nullptr);
	if (FAILED(res))
		throw HRError("Failed to initialize audio client", res);
}

void WASAPISource::InitRender()
{
	CoTaskMemPtr<WAVEFORMATEX> wfex;
	HRESULT                    res;
	LPBYTE                     buffer;
	UINT32                     frames;

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

	res = clientRender->GetMixFormat(&wfex);
	if (FAILED(res))
		throw HRError("Failed to get mix format", res);

	if (isInputDevice) {
		res = clientRender->Initialize(
				AUDCLNT_SHAREMODE_SHARED, AUDCLNT_STREAMFLAGS_LOOPBACK,
				BUFFER_TIME_100NS, 0, wfex, nullptr);
	}
	else {
		res = clientRender->Initialize(
				AUDCLNT_SHAREMODE_SHARED, 0,
				BUFFER_TIME_100NS, 0, wfex, nullptr);
	}

	if (FAILED(res))
		throw HRError("Failed to initialize audio client", res);

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

	memset(buffer, 0, frames*wfex->nBlockAlign);

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

			mt.majortype = MEDIATYPE_Audio;
			mt.subtype = MEDIASUBTYPE_PCM;
			mt.lSampleSize = 0;
			mt.bFixedSizeSamples = TRUE;
			mt.bTemporalCompression = FALSE;
			mt.formattype = FORMAT_WaveFormatEx;
			mt.cbFormat = sizeof(WAVEFORMATEX);

			CoTaskMemPtr<WAVEFORMATEX> wfex;

			res = client->GetMixFormat(&wfex);
			if (FAILED(res))
				throw HRError("Failed to get mix format 0", res);
			nChannels = wfex->nChannels;
			nSampleRate = wfex->nSamplesPerSec;
			mt.pbFormat = (BYTE*)(WAVEFORMATEX*)wfex;
			wfex->wFormatTag = WAVE_FORMAT_PCM;
			wfex->nChannels = 1;
			wfex->nAvgBytesPerSec = nSampleRate * 2;
			wfex->nBlockAlign = 2;
			wfex->wBitsPerSample = 16;
			wfex->cbSize = 0;
			res = captureDMO->SetInputType(0, &mt, 0);
			if (FAILED(res))
				throw HRError("Failed to set input type 0", res);

			res = clientRender->GetMixFormat(&wfex);
			if (FAILED(res))
				throw HRError("Failed to get mix format 1", res);
			nChannelsRender = wfex->nChannels;
			nSampleRateRender = wfex->nSamplesPerSec;
			mt.pbFormat = (BYTE*)(WAVEFORMATEX*)wfex;
			wfex->wFormatTag = WAVE_FORMAT_PCM;
			wfex->nChannels = 1;
			wfex->nAvgBytesPerSec = nSampleRateRender * 2;
			wfex->nBlockAlign = 2;
			wfex->wBitsPerSample = 16;
			wfex->cbSize = 0;
			res = captureDMO->SetInputType(1, &mt, 0);
			if (FAILED(res))
				throw HRError("Failed to set input type 1", res);

			res = MoInitMediaType(&mt, sizeof(WAVEFORMATEX));
			if (FAILED(res))
				throw HRError("Failed to init dmo media type", res);

			WAVEFORMATEX* pwav = (WAVEFORMATEX*)mt.pbFormat;
			pwav->wFormatTag = WAVE_FORMAT_PCM;
			pwav->nChannels = 1;
			pwav->nSamplesPerSec = 22050;
			pwav->nAvgBytesPerSec = 22050 * 2;
			pwav->nBlockAlign = 2;
			pwav->wBitsPerSample = 16;
			pwav->cbSize = 0;

			res = captureDMO->SetOutputType(0, &mt, 0);
			MoFreeMediaType(&mt);

			if (FAILED(res))
				throw HRError("Failed to set dmo output type", res);

			res = captureDMO->AllocateStreamingResources();
			if (FAILED(res))
				throw HRError("Failed to allocate dmo streaming resource", res);

			res = CMediaBuffer::Create(2 * 2 * 22050, captureDMOBuffer.Assign());
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
	InitRender();
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

bool WASAPISource::ProcessCaptureData(bool& dmoActive, FILE* pcmDumpInput, FILE* pcmDumpLoopback, FILE* pcmDumpOutput)
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

		obs_source_audio data = {};
		bool outputDone = false;

		if (captureDMO.Get() && frames) {
			do {
				ComPtr<IMediaBuffer> inMediaBuffer, outMediaBuffer;
				BYTE* pData;

				res = CMediaBuffer::Create(frames * 2, inMediaBuffer.Assign());
				if (FAILED(res)) {
					blog(LOG_ERROR, "Failed to allocate inMediaBuffer %x", res);
					break;
				}
				inMediaBuffer->SetLength(frames * 2);
				inMediaBuffer->GetBufferAndLength(&pData, NULL);
				for (UINT32 i = 0; i < frames; i++) {
					float f = ((float*)buffer)[i * nChannels];
					((short*)pData)[i] = (short)(((f > 1.0f) ? 1.0f : (f < -1.0f) ? -1.0f : f) * 0x7fff);
				}

				res = captureRender->GetNextPacketSize(&captureSize);
				if (FAILED(res) || !captureSize)
					break;

				LPBYTE outbuffer;
				UINT32 outframes = 0;
				UINT64 outPos, outTs;

				res = captureRender->GetBuffer(&outbuffer, &outframes, &flags, &outPos, &outTs);
				if (FAILED(res))
					break;

				if (!outframes) {
					captureRender->ReleaseBuffer(outframes);
					break;
				}

				if (pcmDumpInput != NULL)
					fwrite(pData, 1, frames * 2, pcmDumpInput);

				res = CMediaBuffer::Create(outframes * 2, outMediaBuffer.Assign());
				if (FAILED(res)) {
					blog(LOG_ERROR, "Failed to allocate outMediaBuffer %x", res);
					captureRender->ReleaseBuffer(outframes);
					break;
				}
				outMediaBuffer->SetLength(outframes * 2);
				outMediaBuffer->GetBufferAndLength(&pData, NULL);
				for (UINT32 i = 0; i < outframes; i++) {
					float f = ((float*)outbuffer)[i * nChannelsRender];
					((short*)pData)[i] = (short)(((f > 1.0f) ? 1.0f : (f < -1.0f) ? -1.0f : f) * 0x7fff);
				}

				captureRender->ReleaseBuffer(outframes);

				if (pcmDumpLoopback != NULL)
					fwrite(pData, 1, outframes * 2, pcmDumpLoopback);

				DMO_OUTPUT_DATA_BUFFER dmoOut;
				DWORD dmoStatus;
				dmoOut.pBuffer = captureDMOBuffer;
				captureDMOBuffer->SetLength(0);

				if (!dmoActive) {
					captureDMO->Flush();
					blog(LOG_INFO, "DMO Flush");
				}

				res = captureDMO->ProcessInput(0, inMediaBuffer,
					DMO_INPUT_DATA_BUFFERF_SYNCPOINT|DMO_INPUT_DATA_BUFFERF_TIME|DMO_INPUT_DATA_BUFFERF_TIMELENGTH,
					ts, (10000000ll * frames) / nSampleRate);
				if (FAILED(res)) {
					blog(LOG_ERROR, "Failed to process dmo input 0 %x", res);
					break;
				}

				res = captureDMO->ProcessInput(1, outMediaBuffer,
					DMO_INPUT_DATA_BUFFERF_SYNCPOINT|DMO_INPUT_DATA_BUFFERF_TIME|DMO_INPUT_DATA_BUFFERF_TIMELENGTH,
					outTs, (10000000ll * outframes) / nSampleRateRender);
				if (FAILED(res)) {
					blog(LOG_ERROR, "Failed to process dmo input 1 %x", res);
					break;
				}

				res = captureDMO->ProcessOutput(0, 1, &dmoOut, &dmoStatus);
				if (FAILED(res)) {
					blog(LOG_ERROR, "Failed to process dmo output %x", res);
					break;
				}

				dmoActive = true;
				outputDone = true;

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

					if (pcmDumpOutput != NULL)
						fwrite(pData, 1, len, pcmDumpOutput);
				}
			}
			while (false);
		}

		if (!outputDone) {
			dmoActive = false;

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
	bool dmoActive = true;

#if 0 // for debugging only

	if (source->captureDMO.Get()) {
		pcmDumpInput    = fopen("aec_in0.pcm", "wb");
		pcmDumpLoopback = fopen("aec_in1.pcm", "wb");
		pcmDumpOutput   = fopen("aec_out.pcm", "wb");
	}

#endif

	while (WaitForCaptureSignal(2, sigs, dur)) {
		if (!source->ProcessCaptureData(dmoActive, pcmDumpInput, pcmDumpLoopback, pcmDumpOutput)) {
			reconnect = true;
			break;
		}
	}

	if (pcmDumpInput    != NULL) { fclose(pcmDumpInput); }
	if (pcmDumpLoopback != NULL) { fclose(pcmDumpLoopback); }
	if (pcmDumpOutput   != NULL) { fclose(pcmDumpOutput); }

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
}

static void GetWASAPIDefaultsOutput(obs_data_t *settings)
{
	obs_data_set_default_string(settings, OPT_DEVICE_ID, "default");
	obs_data_set_default_bool(settings, OPT_USE_DEVICE_TIMING, true);
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
