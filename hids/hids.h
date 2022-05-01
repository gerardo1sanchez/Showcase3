
#include "com_smart_pointer.h"
#include <mmreg.h>
#include <dshow.h>

namespace hids
{
	class PropertyBag: public csp::Base<IPropertyBag>
	{
	public:
		csp::Result Read(LPCOLESTR pszPropName, VARIANT &pVar, IErrorLog *pErrorLog);
		csp::Result Write(LPCOLESTR pszPropName, VARIANT &pVar);
		friend class Moniker;
		GENERIC_DEFINITIONS(PropertyBag)
	};

	class PinInfo;

	class KsPropertySet : public csp::Base<IKsPropertySet>
	{
	public:
		csp::Result Set(REFGUID guidPropSet, DWORD dwPropID, LPVOID pInstanceData, DWORD cbInstanceData, LPVOID pPropData, DWORD cbPropData);
		csp::Result Get(REFGUID guidPropSet, DWORD dwPropID, LPVOID pInstanceData, DWORD cbInstanceData, LPVOID pPropData, DWORD cbPropData, DWORD *pcbReturned);
		csp::Result QuerySupported(REFGUID guidPropSet, DWORD dwPropID, DWORD *pTypeSupport);
		friend class Pin;
		GENERIC_DEFINITIONS(KsPropertySet)
	};

	class Pin : public csp::Base<IPin>
	{
	public:
		csp::Result QueryPinInfo(PinInfo& pi);
		CSP_QUERY_INTERFACE(KsPropertySet)
		Pin ConnectedTo();

		friend class CaptureGraphBuilder2;
		friend class GraphBuilder;
		friend class BaseFilter;
		GENERIC_DEFINITIONS(Pin)

	protected:
		CComPtr<IPin>& getObj() { return obj; }
	};

	class EnumPins : public csp::Collection<Pin, IEnumPins>
	{
	public:
		friend class BaseFilter;
		GENERIC_DEFINITIONS(EnumPins)
	};

	class BaseFilter : public csp::Base<IBaseFilter>
	{
	public:
		EnumPins Pins();

		friend class Moniker;
		friend class GraphBuilder;
		friend class Pin;
		friend class CaptureGraphBuilder2;
		GENERIC_DEFINITIONS(BaseFilter)
	};

	class PinInfo
	{
	public:
		PinInfo() {}
		PinInfo(const PinInfo& cpy);

		BaseFilter	  pFilter;
		PIN_DIRECTION dir;
		WCHAR         achName[MAX_PIN_NAME];
	};

	class Moniker: public csp::Base<IMoniker>
	{
	public:
		csp::Result CreateItemMoniker(LPCOLESTR lpszDelim, LPCOLESTR lpszItem);
		BaseFilter BindToObjectBaseFilter();
		bool operator !=(const Moniker& right) const;
		bool operator ==(const Moniker& right) const;

		PropertyBag BindToStoragePropertyBag();

		friend class RunningObjectTable;
		friend class csp::Collection<Moniker, IEnumMoniker>;
		friend class csp::iterator<Moniker, IEnumMoniker>;
		GENERIC_DEFINITIONS(Moniker)
	};

	class EnumMoniker: public csp::Collection<Moniker, IEnumMoniker>
	{
	public:
		friend class SystemDeviceEnum;
		friend class RunningObjectTable;
		GENERIC_DEFINITIONS(EnumMoniker)
	};

	class SystemDeviceEnum : public csp::Base<ICreateDevEnum>
	{
	public:
		csp::Result CreateInstance();
		EnumMoniker CreateClassEnumerator(REFCLSID clsidDeviceClass, DWORD dwFlags);
		GENERIC_DEFINITIONS(SystemDeviceEnum)
	};

	class MediaControl : public csp::Base<IMediaControl>
	{
	public:
		csp::Result Run();
		csp::Result Pause();
		csp::Result Stop();
		csp::Result GetState(LONG msTimeout, OAFilterState &pfs);
		csp::Result RenderFile(BSTR strFilename);
		csp::Result StopWhenReady();

		friend class GraphBuilder;
		GENERIC_DEFINITIONS(MediaControl)
	};

	class BasicVideo: public csp::Base<IBasicVideo>
	{
	public:
		CSP_PROP_1(SourceLeft, long)
		CSP_PROP_1(SourceTop, long)
		CSP_PROP_1(SourceWidth, long)
		CSP_PROP_1(SourceHeight, long)
		CSP_PROP_1(DestinationLeft, long)
		CSP_PROP_1(DestinationTop, long)
		CSP_PROP_1(DestinationWidth, long)
		CSP_PROP_1(DestinationHeight, long)

		csp::Result get_AvgTimePerFrame(REFTIME &pAvgTimePerFrame);
		csp::Result get_BitRate(long &pBitRate);
		csp::Result get_BitErrorRate(long &pBitErrorRate);
		csp::Result get_VideoWidth(long &pVideoWidth);
		csp::Result get_VideoHeight(long &pVideoHeight);
		csp::Result SetSourcePosition(long Left, long Top, long Width, long Height);
		csp::Result GetSourcePosition(long &pLeft, long &pTop, long &pWidth, long &pHeight);
		csp::Result SetDefaultSourcePosition(void);
		csp::Result SetDestinationPosition(long Left, long Top, long Width, long Height);
		csp::Result GetDestinationPosition(long &pLeft, long &pTop, long &pWidth, long &pHeight);
		csp::Result SetDefaultDestinationPosition();
		csp::Result GetVideoSize(long &pWidth,long &pHeight);
		csp::Result GetVideoPaletteEntries(long StartIndex, long Entries, long *pRetrieved, long *pPalette);
		csp::Result GetCurrentImage(long *pBufferSize, long *pDIBImage);
		csp::Result IsUsingDefaultSource();
		csp::Result IsUsingDefaultDestination();

		friend class GraphBuilder;
		GENERIC_DEFINITIONS(BasicVideo)
	};
	
	class VideoWindow : public csp::Base<IVideoWindow>
	{
	public:
		CSP_PROP_1(WindowStyle, long)
		CSP_PROP_1(WindowStyleEx, long)
		CSP_PROP_1(AutoShow, long)
		CSP_PROP_1(WindowState, long)
		CSP_PROP_1(BackgroundPalette, long)
		CSP_PROP_1(Visible, long)
		CSP_PROP_1(Left, long)
		CSP_PROP_1(Width, long)
		CSP_PROP_1(Top, long)
		CSP_PROP_1(Height, long)
		CSP_PROP_1(Owner, OAHWND)
		CSP_PROP_1(MessageDrain, OAHWND)
		CSP_PROP_1(BorderColor, long)
		CSP_PROP_1(FullScreenMode, long)

		csp::Result put_Caption(BSTR strCaption);
		csp::Result get_Caption(BSTR *strCaption);

		csp::Result SetWindowPosition(long Left, long Top, long Width, long Height);
		csp::Result GetWindowPosition(long *pLeft, long *pTop, long *pWidth, long *pHeight);

		csp::Result SetWindowForeground(long Focus);
		csp::Result NotifyOwnerMessage(OAHWND hwnd, long uMsg, LONG_PTR wParam, LONG_PTR lParam);
		csp::Result GetMinIdealImageSize(long & pWidth, long & pHeight);
		csp::Result GetMaxIdealImageSize(long &pWidth, long &pHeight);
		csp::Result GetRestorePosition(long *pLeft, long *pTop, long *pWidth, long *pHeight);
		csp::Result HideCursor(long HideCursor);
		csp::Result IsCursorHidden(long &CursorHidden);

		friend class GraphBuilder;
		GENERIC_DEFINITIONS(VideoWindow)
	};

	class MediaEventEx : public csp::Base<IMediaEventEx>
	{
	public:
		csp::Result GetEventHandle(OAEVENT &hEvent);
		csp::Result GetEvent(long &lEventCode, LONG_PTR *lParam1, LONG_PTR *lParam2, long msTimeout);
		csp::Result WaitForCompletion(long msTimeout, long &pEvCode);
		csp::Result CancelDefaultHandling(long lEvCode);
		csp::Result RestoreDefaultHandling(long lEvCode);
		csp::Result FreeEventParams(long lEvCode, LONG_PTR lParam1, LONG_PTR lParam2);

		csp::Result SetNotifyWindow(OAHWND hwnd, long lMsg, LONG_PTR lInstanceData);

		CSP_PROP_2(NotifyFlags, long)

		friend class GraphBuilder;
		GENERIC_DEFINITIONS(MediaEventEx)
	};

	class GraphBuilder : public csp::Base<IGraphBuilder>
	{
	public:
		csp::Result CreateInstance();
		csp::Result AddFilter(BaseFilter& filter, LPCWSTR pName);
		csp::Result RemoveFilter(BaseFilter &pFilter);
		CSP_QUERY_INTERFACE(MediaControl)
		CSP_QUERY_INTERFACE(VideoWindow)
		CSP_QUERY_INTERFACE(MediaEventEx)
		CSP_QUERY_INTERFACE(BasicVideo)
		csp::Result Disconnect(Pin &ppin);

		friend class CaptureGraphBuilder2;
		GENERIC_DEFINITIONS(GraphBuilder)
	};

	class FileSinkFilter: public csp::Base<IFileSinkFilter>
	{
	public:
		csp::Result SetFileName(LPCOLESTR pszFileName, const AM_MEDIA_TYPE &pmt);
		csp::Result GetCurFile(LPOLESTR *ppszFileName, AM_MEDIA_TYPE *pmt);
		friend class CaptureGraphBuilder2;
		GENERIC_DEFINITIONS(FileSinkFilter)
	};

	class CaptureGraphBuilder2 : public csp::Base<ICaptureGraphBuilder2>
	{
	public:
		csp::Result CreateInstance();
		csp::Result SetFiltergraph(GraphBuilder &pfg);
		GraphBuilder GetFiltergraph();
		csp::Result SetOutputFileName(const GUID *pType, LPCOLESTR lpstrFile, BaseFilter &ppf, FileSinkFilter *ppSink = nullptr);
		csp::Result RenderStream(const GUID *pCategory, const GUID *pType, BaseFilter &pSource, BaseFilter *pfCompressor, BaseFilter *pfRenderer);
		csp::Result RenderStream(const GUID *pCategory, const GUID *pType, Pin &pSource, BaseFilter *pfCompressor, BaseFilter *pfRenderer);
		csp::Unknown FindInterface(const GUID *pCategory, const GUID *pType, BaseFilter &pf, REFIID riid);
		Pin FindPin(csp::Unknown& pSource, PIN_DIRECTION pindir, const GUID *pCategory, const GUID *pType, BOOL fUnconnected, int num);

		GENERIC_DEFINITIONS(CaptureGraphBuilder2)
	};

	class AmMediaTypeHolder
	{
	public:
		AmMediaTypeHolder();
		~AmMediaTypeHolder();
		VIDEOINFOHEADER* GetVideoInfoHeader() const;

	private:
		AmMediaTypeHolder(const AmMediaTypeHolder&) {}
		void operator=(const AmMediaTypeHolder&) {}
		AM_MEDIA_TYPE *media;
		friend class AMStreamConfig;
	};

	class AMStreamConfig : public csp::Base<IAMStreamConfig>
	{
	public:
		csp::Result GetFormat(AmMediaTypeHolder& amt);
		GENERIC_DEFINITIONS(AMStreamConfig)
	};

	class AMVideoCompression : public csp::Base<IAMVideoCompression>
	{
	public:
		CSP_PROP_1(KeyFrameRate, long)
		CSP_PROP_1(PFramesPerKeyFrame, long)
		CSP_PROP_1(Quality, double)
		CSP_PROP_1(WindowSize, DWORDLONG)

		csp::Result GetInfo(LPWSTR pszVersion, int *pcbVersion, LPWSTR pszDescription, int *pcbDescription, long *pDefaultKeyFrameRate, long *pDefaultPFramesPerKey, double *pDefaultQuality, long *pCapabilities);
		csp::Result OverrideKeyFrame(long FrameNumber);
		csp::Result OverrideFrameSize(long FrameNumber, long Size);

		GENERIC_DEFINITIONS(AMVideoCompression)
	};

	class RunningObjectTable: public csp::Base<IRunningObjectTable>
	{
	public:
		csp::Result CreateInstance();
		csp::Result Register(DWORD grfFlags, const csp::Unknown &punkObject, const Moniker &pmkObjectName, DWORD &pdwRegister);
		csp::Result Revoke(DWORD dwRegister);
		csp::Result IsRunning(const Moniker &pmkObjectName);
		csp::Unknown GetObject(Moniker &pmkObjectName);
		csp::Result NoteChangeTime(DWORD dwRegister, FILETIME &pfiletime);
		csp::Result GetTimeOfLastChange(const Moniker &pmkObjectName, FILETIME &pfiletime);
		EnumMoniker EnumRunning();

		GENERIC_DEFINITIONS(RunningObjectTable)
	};

}

#define Trace2(context_class, context_method, message)				/*{ const type_info &__i = typeid(*this); LPCSTR name = __i.name(); TraceObject.WriteMsg(__FILE__, __LINE__, name?name:#context_class, #context_method, #message, (LPCTSTR)NULL); }*/
#define CheckArguments(cond, context_class, context_method)	{if(cond) {Trace2(context_class, context_method, "inval. argments"); DebugInt3; return E_INVALIDARG;}}
#define CheckInitialized(context_class, context_method)		{if(!IsValid()) {Trace2(context_class, context_method, "!init"); DebugInt3; return E_HANDLE;}}

