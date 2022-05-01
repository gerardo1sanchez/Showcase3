
#include <windows.h>
#include <atlcomcli.h>

// clonar https://github.com/microsoft/Windows-classic-samples.git,
// y compilar ...Windows-classic-samples\Samples\Win7Samples\multimedia\directshow\baseclasses\baseclasses.sln
#include <dshowutil.h>
#include "hids.h"

using namespace hids;

PropertyBag Moniker::BindToStoragePropertyBag()
{
	if (!obj) return PropertyBag();
	PropertyBag pb;
	if (SUCCEEDED(obj->BindToStorage(0, 0, IID_IPropertyBag, (void **)&pb.obj))) return pb;
	return PropertyBag();
}

csp::Result SystemDeviceEnum::CreateInstance()
{
	if (obj) return NOERROR;
	return obj.CoCreateInstance(CLSID_SystemDeviceEnum, nullptr, CLSCTX_INPROC_SERVER);
}

EnumMoniker SystemDeviceEnum::CreateClassEnumerator(REFCLSID clsidDeviceClass, DWORD dwFlags)
{
	if (!CreateInstance()) return EnumMoniker();

	EnumMoniker pEm;
	if (SUCCEEDED(obj->CreateClassEnumerator(CLSID_VideoInputDeviceCategory, &pEm.obj, 0))) return pEm;
	return EnumMoniker();
}

csp::Result Moniker::CreateItemMoniker(LPCOLESTR lpszDelim, LPCOLESTR lpszItem)
{
	obj.Release();
	return ::CreateItemMoniker(lpszDelim, lpszItem, &obj);
}

BaseFilter Moniker::BindToObjectBaseFilter()
{
	BaseFilter bf;
	if (!IsValid()) return bf;
	if(SUCCEEDED(obj->BindToObject(0, 0, IID_IBaseFilter, (void**)&bf.obj))) return bf;
	return BaseFilter();
}

bool Moniker::operator !=(const Moniker& right) const
{
	return obj != right.obj;

}

bool Moniker::operator ==(const Moniker& right) const
{
	return obj == right.obj;
}

csp::Result GraphBuilder::CreateInstance()
{
	if (obj) return NOERROR;
	return obj.CoCreateInstance(CLSID_FilterGraph/* yes, a CLSID_FilterGraph returns a IGraphBuilder instance*/, nullptr, CLSCTX_INPROC_SERVER);
}

csp::Result GraphBuilder::AddFilter(BaseFilter& filter, LPCWSTR pName)
{
	CheckArguments(!filter.IsValid() || !pName, GraphBuilder, AddFilter);
	HRESULT hr = CreateInstance().last_error;
	if (!SUCCEEDED(hr)) return hr;
	return obj->AddFilter(filter.obj, pName);
}

csp::Result GraphBuilder::RemoveFilter(BaseFilter & filter)
{
	CheckArguments(!filter.IsValid(), GraphBuilder, RemoveFilter);
	CheckInitialized(GraphBuilder, RemoveFilter);
	return SUCCEEDED(obj->RemoveFilter(filter.obj));
}

CSP_QUERY_INTERFACE_IMP(GraphBuilder, MediaControl)
CSP_QUERY_INTERFACE_IMP(GraphBuilder, VideoWindow)
CSP_QUERY_INTERFACE_IMP(GraphBuilder, MediaEventEx)
CSP_QUERY_INTERFACE_IMP(GraphBuilder, BasicVideo)

csp::Result GraphBuilder::Disconnect(Pin & ppin)
{
	CheckArguments(!ppin.IsValid(), GraphBuilder, Disconnect);
	CheckInitialized(GraphBuilder, Disconnect);
	return obj->Disconnect(ppin.obj);
}

csp::Result MediaControl::Run()
{
	CheckInitialized(MediaControl, Run);
	return obj->Run();
}

csp::Result MediaControl::Pause()
{
	CheckInitialized(MediaControl, Pause);
	return obj->Pause();
}

csp::Result MediaControl::Stop()
{
	CheckInitialized(MediaControl, Stop);
	return obj->Stop();
}

csp::Result MediaControl::GetState(LONG msTimeout, OAFilterState &pfs)
{
	CheckInitialized(MediaControl, GetState);
	return obj->GetState(msTimeout, &pfs);
}

csp::Result MediaControl::RenderFile(BSTR strFilename)
{
	CheckArguments(!strFilename, MediaControl, RenderFile);
	CheckInitialized(MediaControl, RenderFile);
	return obj->RenderFile(strFilename);
}

csp::Result MediaControl::StopWhenReady()
{
	CheckInitialized(MediaControl, StopWhenReady);
	return obj->StopWhenReady();
}

csp::Result AMStreamConfig::GetFormat(AmMediaTypeHolder& amt)
{
	CheckInitialized(AMStreamConfig, GetFormat);
	if (amt.media)
	{
		DeleteMediaType(amt.media);
		amt.media = nullptr;
	}
	return obj->GetFormat(&amt.media);
}

AmMediaTypeHolder::AmMediaTypeHolder()
{
	media = nullptr;
}

AmMediaTypeHolder::~AmMediaTypeHolder()
{
	if(media) DeleteMediaType(media);
}

VIDEOINFOHEADER * AmMediaTypeHolder::GetVideoInfoHeader() const
{
	if(media != nullptr && media->majortype == MEDIATYPE_Video && media->cbFormat == sizeof(VIDEOINFOHEADER)) return (VIDEOINFOHEADER*)media->pbFormat;
	return nullptr;
}

csp::Result hids::VideoWindow::SetWindowForeground(long Focus)
{
	CheckInitialized(VideoWindow, get_Caption);
	return obj->SetWindowForeground(Focus);
}

csp::Result hids::VideoWindow::NotifyOwnerMessage(OAHWND hwnd, long uMsg, LONG_PTR wParam, LONG_PTR lParam)
{
	CheckInitialized(VideoWindow, NotifyOwnerMessage);
	return obj->NotifyOwnerMessage(hwnd, uMsg, wParam, lParam);
}

csp::Result VideoWindow::get_Caption(BSTR * strCaption)
{
	CheckInitialized(VideoWindow, get_Caption);
	return obj->get_Caption(strCaption);
}

csp::Result VideoWindow::SetWindowPosition(long Left, long Top, long Width, long Height)
{
	CheckInitialized(VideoWindow, SetWindowPosition);
	return obj->SetWindowPosition(Left, Top, Width, Height);
}

csp::Result VideoWindow::GetWindowPosition(long * pLeft, long * pTop, long * pWidth, long * pHeight)
{
	CheckInitialized(VideoWindow, GetWindowPosition);
	return obj->GetWindowPosition(pLeft, pTop, pWidth, pHeight);
}

CSP_PROP_1_IMP(VideoWindow, WindowStyle, long)
CSP_PROP_1_IMP(VideoWindow, WindowStyleEx, long)
CSP_PROP_1_IMP(VideoWindow, AutoShow, long)
CSP_PROP_1_IMP(VideoWindow, WindowState, long)
CSP_PROP_1_IMP(VideoWindow, BackgroundPalette, long)
CSP_PROP_1_IMP(VideoWindow, Visible, long)
CSP_PROP_1_IMP(VideoWindow, Left, long)
CSP_PROP_1_IMP(VideoWindow, Width, long)
CSP_PROP_1_IMP(VideoWindow, Top, long)
CSP_PROP_1_IMP(VideoWindow, Height, long)
CSP_PROP_1_IMP(VideoWindow, Owner, OAHWND)
CSP_PROP_1_IMP(VideoWindow, MessageDrain, OAHWND)
CSP_PROP_1_IMP(VideoWindow, BorderColor, long)
CSP_PROP_1_IMP(VideoWindow, FullScreenMode, long);

csp::Result VideoWindow::GetMinIdealImageSize(long & pWidth, long & pHeight)
{
	CheckInitialized(VideoWindow, GetMinIdealImageSize);
	return obj->GetMinIdealImageSize(&pWidth, &pHeight);
}

csp::Result VideoWindow::GetMaxIdealImageSize(long &pWidth, long &pHeight)
{
	CheckInitialized(VideoWindow, GetMaxIdealImageSize);
	return obj->GetMaxIdealImageSize(&pWidth, &pHeight);
}

csp::Result VideoWindow::GetRestorePosition(long * pLeft, long * pTop, long * pWidth, long * pHeight)
{
	CheckInitialized(VideoWindow, GetRestorePosition);
	return obj->GetRestorePosition(pLeft, pTop, pWidth, pHeight);
}

csp::Result VideoWindow::HideCursor(long HideCursor)
{
	CheckInitialized(VideoWindow, IsCursorHidden);
	return obj->HideCursor(HideCursor);
}

csp::Result VideoWindow::IsCursorHidden(long &CursorHidden)
{
	CheckInitialized(VideoWindow, IsCursorHidden);
	return obj->IsCursorHidden(&CursorHidden);
}

csp::Result PropertyBag::Read(LPCOLESTR pszPropName, VARIANT &pVar, IErrorLog * pErrorLog)
{
	CheckArguments(!pszPropName, PropertyBag, Read);
	CheckInitialized(PropertyBag, Read);
	return obj->Read(pszPropName, &pVar, pErrorLog);
}

csp::Result PropertyBag::Write(LPCOLESTR pszPropName, VARIANT &pVar)
{
	CheckArguments(!pszPropName, PropertyBag, Write);
	CheckInitialized(PropertyBag, Write);
	return obj->Write(pszPropName, &pVar);
}

EnumPins BaseFilter::Pins()
{
	EnumPins p;
	if (!IsValid()) return p;
	if (SUCCEEDED(obj->EnumPins(&p.obj))) return p;
	return EnumPins();
}

csp::Result Pin::QueryPinInfo(PinInfo & pi)
{
	CheckInitialized(Pin, QueryPinInfo);
	PIN_INFO pinInfo;
	HRESULT hr = obj->QueryPinInfo(&pinInfo);
	if(!SUCCEEDED(hr)) return hr;
	pi.pFilter.obj.Release();
	pi.pFilter.obj.Attach(pinInfo.pFilter);
	pi.dir = pinInfo.dir;
	lstrcpyW(pi.achName, pinInfo.achName);
	return hr;
}

CSP_QUERY_INTERFACE_IMP(Pin, KsPropertySet)

Pin Pin::ConnectedTo()
{
	Pin p;
	if (!IsValid()) return p;
	if (SUCCEEDED(obj->ConnectedTo(&p.obj))) return p;
	return Pin();
}

PinInfo::PinInfo(const PinInfo & cpy)
{
	pFilter = cpy.pFilter;
	dir = cpy.dir;
	lstrcpyW(achName, cpy.achName);
}

csp::Result KsPropertySet::Set(REFGUID guidPropSet, DWORD dwPropID, LPVOID pInstanceData, DWORD cbInstanceData, LPVOID pPropData, DWORD cbPropData)
{
	CheckInitialized(KsPropertySet, Set);
	return obj->Set(guidPropSet, dwPropID, pInstanceData, cbInstanceData, pPropData, cbPropData);
}

csp::Result KsPropertySet::Get(REFGUID guidPropSet, DWORD dwPropID, LPVOID pInstanceData, DWORD cbInstanceData, LPVOID pPropData, DWORD cbPropData, DWORD *pcbReturned)
{
	CheckInitialized(KsPropertySet, Get);
	return obj->Get(guidPropSet, dwPropID, pInstanceData, cbInstanceData, pPropData, cbPropData, pcbReturned);
}

csp::Result KsPropertySet::QuerySupported(REFGUID guidPropSet, DWORD dwPropID, DWORD * pTypeSupport)
{
	CheckInitialized(KsPropertySet, QuerySupported);
	return obj->QuerySupported(guidPropSet, dwPropID, pTypeSupport);
}

csp::Result CaptureGraphBuilder2::CreateInstance()
{
	if (obj) return NOERROR;
	return obj.CoCreateInstance(CLSID_CaptureGraphBuilder2, nullptr, CLSCTX_INPROC_SERVER);
}

csp::Result CaptureGraphBuilder2::SetFiltergraph(GraphBuilder & pfg)
{
	HRESULT hr = CreateInstance().last_error;
	if (!SUCCEEDED(hr)) return hr;
	CheckArguments(!pfg, CaptureGraphBuilder2, SetFiltergraph);
	return obj->SetFiltergraph(pfg.obj);
}

GraphBuilder CaptureGraphBuilder2::GetFiltergraph()
{
	GraphBuilder p;
	if (!CreateInstance()) return p;
	if (SUCCEEDED(obj->GetFiltergraph(&p.obj))) return p;
	return GraphBuilder();
}

csp::Result CaptureGraphBuilder2::SetOutputFileName(const GUID * pType, LPCOLESTR lpstrFile, BaseFilter & ppf, FileSinkFilter * ppSink)
{
	HRESULT hr = CreateInstance().last_error;
	if (!SUCCEEDED(hr)) return hr;
	ppf.Release();
	if (ppSink) ppSink->Release();
	return obj->SetOutputFileName(pType, lpstrFile, &ppf.obj, ppSink ?(&ppSink->obj):nullptr);
}

csp::Result CaptureGraphBuilder2::RenderStream(const GUID * pCategory, const GUID * pType, BaseFilter & pSource, BaseFilter * pfCompressor, BaseFilter * pfRenderer)
{
	HRESULT hr = CreateInstance().last_error;
	if (!SUCCEEDED(hr)) return hr;
	CheckArguments(!pSource || (pfCompressor != nullptr && !pfCompressor->obj) || (pfRenderer != nullptr && !pfRenderer->obj), FileSinkFilter, RenderStream);
	return obj->RenderStream(pCategory, pType, pSource.obj, pfCompressor ? (pfCompressor->obj) : nullptr, pfRenderer ? (pfRenderer->obj) : nullptr);
}

csp::Result CaptureGraphBuilder2::RenderStream(const GUID * pCategory, const GUID * pType, Pin & pSource, BaseFilter * pfCompressor, BaseFilter * pfRenderer)
{
	HRESULT hr = CreateInstance().last_error;
	if (!SUCCEEDED(hr)) return hr;
	CheckArguments(!pSource || (pfCompressor != nullptr && !pfCompressor->obj) || (pfRenderer != nullptr && !pfRenderer->obj), FileSinkFilter, RenderStream);
	return obj->RenderStream(pCategory, pType, pSource.obj, pfCompressor?(pfCompressor->obj):nullptr, pfRenderer?(pfRenderer->obj):nullptr);
}

csp::Unknown CaptureGraphBuilder2::FindInterface(const GUID * pCategory, const GUID * pType, BaseFilter & pf, REFIID riid)
{
	if (!IsValid()) return csp::Unknown();
	CComPtr<IUnknown> p;
	if (SUCCEEDED(obj->FindInterface(pCategory, pType, pf.obj, riid, (void**)&p))) return p;
	return csp::Unknown(p);
}

Pin CaptureGraphBuilder2::FindPin(csp::Unknown & pSource, PIN_DIRECTION pindir, const GUID * pCategory, const GUID * pType, BOOL fUnconnected, int num)
{
	Pin mc;
	if (!pSource || !IsValid()) return mc;
	if (SUCCEEDED(obj->FindPin(pSource.getObj(), pindir, pCategory, pType, fUnconnected, num, &mc.getObj()))) return mc;
	return Pin();
}

csp::Result FileSinkFilter::SetFileName(LPCOLESTR pszFileName, const AM_MEDIA_TYPE &pmt)
{
	CheckInitialized(FileSinkFilter, SetFileName);
	CheckArguments(!pszFileName, FileSinkFilter, SetFileName);
	return obj->SetFileName(pszFileName, &pmt);
}

csp::Result FileSinkFilter::GetCurFile(LPOLESTR * ppszFileName, AM_MEDIA_TYPE *pmt)
{
	CheckInitialized(FileSinkFilter, GetCurFile);
	CheckArguments(!ppszFileName || !pmt, FileSinkFilter, GetCurFile);
	return obj->GetCurFile(ppszFileName, pmt);
}

CSP_PROP_1_IMP(AMVideoCompression, KeyFrameRate, long)
CSP_PROP_1_IMP(AMVideoCompression, PFramesPerKeyFrame, long)
CSP_PROP_1_IMP(AMVideoCompression, Quality, double)
CSP_PROP_1_IMP(AMVideoCompression, WindowSize, DWORDLONG)

csp::Result AMVideoCompression::GetInfo(LPWSTR pszVersion, int * pcbVersion, LPWSTR pszDescription, int * pcbDescription, long * pDefaultKeyFrameRate, long * pDefaultPFramesPerKey, double * pDefaultQuality, long * pCapabilities)
{
	CheckInitialized(AMVideoCompression, GetInfo);
	return obj->GetInfo(pszVersion, pcbVersion, pszDescription, pcbDescription, pDefaultKeyFrameRate, pDefaultPFramesPerKey, pDefaultQuality, pCapabilities);
}

csp::Result AMVideoCompression::OverrideKeyFrame(long FrameNumber)
{
	CheckInitialized(AMVideoCompression, OverrideKeyFrame);
	return obj->OverrideKeyFrame(FrameNumber);
}

csp::Result AMVideoCompression::OverrideFrameSize(long FrameNumber, long Size)
{
	CheckInitialized(AMVideoCompression, OverrideFrameSize);
	return obj->OverrideFrameSize(FrameNumber, Size);
}

csp::Result RunningObjectTable::CreateInstance()
{
	if (obj) return NOERROR;
	return GetRunningObjectTable(0, &obj);
}

csp::Result RunningObjectTable::Register(DWORD grfFlags, const csp::Unknown & punkObject, const Moniker & pmkObjectName, DWORD & pdwRegister)
{
	CheckArguments(!punkObject || !pmkObjectName, RunningObjectTable, Register);
	HRESULT hr = CreateInstance().last_error;
	if (!SUCCEEDED(hr)) return hr;
	return obj->Register(grfFlags, punkObject.getObj(), pmkObjectName.obj, &pdwRegister);
}

csp::Result RunningObjectTable::Revoke(DWORD dwRegister)
{
	HRESULT hr = CreateInstance().last_error;
	if (!SUCCEEDED(hr)) return hr;
	return obj->Revoke(dwRegister);
}

csp::Result RunningObjectTable::IsRunning(const Moniker & pmkObjectName)
{
	CheckArguments(!pmkObjectName, RunningObjectTable, IsRunning);
	HRESULT hr = CreateInstance().last_error;
	if (!SUCCEEDED(hr)) return hr;
	return obj->IsRunning(pmkObjectName.obj);
}

csp::Unknown RunningObjectTable::GetObject(Moniker & pmkObjectName)
{
	if (!pmkObjectName) return csp::Unknown();
	if (!CreateInstance()) return csp::Unknown();
	CComPtr<IUnknown> u;
	HRESULT hr = obj->GetObject(pmkObjectName.obj, &u);
	if (SUCCEEDED(hr)) return csp::Unknown(u);
	return csp::Unknown();
}

csp::Result RunningObjectTable::NoteChangeTime(DWORD dwRegister, FILETIME & pfiletime)
{
	HRESULT hr = CreateInstance().last_error;
	if (!SUCCEEDED(hr)) return hr;
	return obj->NoteChangeTime(dwRegister, &pfiletime);
}

csp::Result RunningObjectTable::GetTimeOfLastChange(const Moniker & pmkObjectName, FILETIME & pfiletime)
{
	HRESULT hr = CreateInstance().last_error;
	if (!SUCCEEDED(hr)) return hr;
	return obj->GetTimeOfLastChange(pmkObjectName.obj, &pfiletime);
}

EnumMoniker RunningObjectTable::EnumRunning()
{
	EnumMoniker e;
	HRESULT hr = CreateInstance().last_error;
	if (!SUCCEEDED(hr)) return e;
	hr = obj->EnumRunning(&e.obj);
	if (SUCCEEDED(hr)) return e;
	return EnumMoniker();
}

csp::Result MediaEventEx::GetEventHandle(OAEVENT &hEvent)
{
	CheckInitialized(MediaEventEx, GetEventHandle);
	return obj->GetEventHandle(&hEvent);
}

csp::Result MediaEventEx::GetEvent(long &lEventCode, LONG_PTR * lParam1, LONG_PTR * lParam2, long msTimeout)
{
	CheckInitialized(MediaEventEx, GetEvent);
	return obj->GetEvent(&lEventCode, lParam1, lParam2, msTimeout);
}

csp::Result MediaEventEx::WaitForCompletion(long msTimeout, long &pEvCode)
{
	CheckInitialized(MediaEventEx, WaitForCompletion);
	return obj->WaitForCompletion(msTimeout, &pEvCode);
}

csp::Result MediaEventEx::CancelDefaultHandling(long lEvCode)
{
	CheckInitialized(MediaEventEx, CancelDefaultHandling);
	return obj->CancelDefaultHandling(lEvCode);
}

csp::Result MediaEventEx::RestoreDefaultHandling(long lEvCode)
{
	CheckInitialized(MediaEventEx, RestoreDefaultHandling);
	return obj->RestoreDefaultHandling(lEvCode);
}

csp::Result MediaEventEx::FreeEventParams(long lEvCode, LONG_PTR lParam1, LONG_PTR lParam2)
{
	CheckInitialized(MediaEventEx, FreeEventParams);
	return obj->FreeEventParams(lEvCode, lParam1, lParam2);
}

csp::Result MediaEventEx::SetNotifyWindow(OAHWND hwnd, long lMsg, LONG_PTR lInstanceData)
{
	CheckInitialized(MediaEventEx, SetNotifyWindow);
	return obj->SetNotifyWindow(hwnd, lMsg, lInstanceData);
}

CSP_PROP_2_IMP(MediaEventEx, NotifyFlags, long)

CSP_PROP_1_IMP(BasicVideo, SourceLeft, long)
CSP_PROP_1_IMP(BasicVideo, SourceTop, long)
CSP_PROP_1_IMP(BasicVideo, SourceWidth, long)
CSP_PROP_1_IMP(BasicVideo, SourceHeight, long)
CSP_PROP_1_IMP(BasicVideo, DestinationLeft, long)
CSP_PROP_1_IMP(BasicVideo, DestinationTop, long)
CSP_PROP_1_IMP(BasicVideo, DestinationWidth, long)
CSP_PROP_1_IMP(BasicVideo, DestinationHeight, long);

csp::Result BasicVideo::get_AvgTimePerFrame(REFTIME &pAvgTimePerFrame)
{
	CheckInitialized(BasicVideo, get_AvgTimePerFrame);
	return obj->get_AvgTimePerFrame(&pAvgTimePerFrame);
}

csp::Result BasicVideo::get_BitRate(long &pBitRate)
{
	CheckInitialized(BasicVideo, get_BitRate);
	return obj->get_BitRate(&pBitRate);
}

csp::Result BasicVideo::get_BitErrorRate(long &pBitErrorRate)
{
	CheckInitialized(BasicVideo, get_BitErrorRate);
	return obj->get_BitErrorRate(&pBitErrorRate);
}

csp::Result BasicVideo::get_VideoWidth(long &pVideoWidth)
{
	CheckInitialized(BasicVideo, get_VideoWidth);
	return obj->get_VideoWidth(&pVideoWidth);
}

csp::Result BasicVideo::get_VideoHeight(long &pVideoHeight)
{
	CheckInitialized(BasicVideo, get_VideoHeight);
	return obj->get_VideoHeight(&pVideoHeight);
}

csp::Result BasicVideo::SetSourcePosition(long Left, long Top, long Width, long Height)
{
	CheckInitialized(BasicVideo, SetSourcePosition);
	return obj->SetSourcePosition(Left, Top, Width, Height);
}

csp::Result BasicVideo::GetSourcePosition(long & pLeft, long & pTop, long & pWidth, long & pHeight)
{
	CheckInitialized(BasicVideo, GetSourcePosition);
	return obj->GetSourcePosition(&pLeft, &pTop, &pWidth, &pHeight);
}

csp::Result BasicVideo::SetDefaultSourcePosition()
{
	CheckInitialized(BasicVideo, SetDefaultSourcePosition);
	return obj->SetDefaultSourcePosition();
}

csp::Result BasicVideo::SetDestinationPosition(long Left, long Top, long Width, long Height)
{
	CheckInitialized(BasicVideo, SetDestinationPosition);
	return obj->SetDestinationPosition(Left, Top, Width, Height);
}

csp::Result BasicVideo::GetDestinationPosition(long & pLeft, long & pTop, long & pWidth, long & pHeight)
{
	CheckInitialized(BasicVideo, GetDestinationPosition);
	return obj->GetDestinationPosition(&pLeft, &pTop, &pWidth, &pHeight);
}

csp::Result BasicVideo::SetDefaultDestinationPosition()
{
	CheckInitialized(BasicVideo, SetDefaultDestinationPosition);
	return obj->SetDefaultDestinationPosition();
}

csp::Result BasicVideo::GetVideoSize(long & pWidth, long & pHeight)
{
	CheckInitialized(BasicVideo, GetVideoSize);
	return obj->GetVideoSize(&pWidth, &pHeight);
}

csp::Result BasicVideo::GetVideoPaletteEntries(long StartIndex, long Entries, long * pRetrieved, long * pPalette)
{
	CheckInitialized(BasicVideo, GetVideoPaletteEntries);
	return obj->GetVideoPaletteEntries(StartIndex, Entries, pRetrieved, pPalette);
}

csp::Result BasicVideo::GetCurrentImage(long * pBufferSize, long * pDIBImage)
{
	CheckInitialized(BasicVideo, GetCurrentImage);
	return obj->GetCurrentImage(pBufferSize, pDIBImage);
}

csp::Result BasicVideo::IsUsingDefaultSource()
{
	CheckInitialized(BasicVideo, IsUsingDefaultSource);
	return obj->IsUsingDefaultSource();
}

csp::Result BasicVideo::IsUsingDefaultDestination()
{
	CheckInitialized(BasicVideo, IsUsingDefaultDestination);
	return obj->IsUsingDefaultDestination();
}
