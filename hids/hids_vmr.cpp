
#include <windows.h>
#include <atlcomcli.h>

// clonar https://github.com/microsoft/Windows-classic-samples.git,
// y compilar ...Windows-classic-samples\Samples\Win7Samples\multimedia\directshow\baseclasses\baseclasses.sln
#include <dshowutil.h>
#include "hids.h"
#include "hids_vmr.h"

using namespace hids;
using namespace hids::vmr9;

Result Renderer::CreateInstance()
{
	if (obj) return NOERROR;
	return obj.CoCreateInstance(CLSID_VideoMixingRenderer9, nullptr, CLSCTX_INPROC);
}

CSP_QUERY_INTERFACE_IMP(hids::vmr9::Renderer, FilterConfig)
CSP_QUERY_INTERFACE_IMP(hids::vmr9::Renderer, WindowlessControl)
CSP_QUERY_INTERFACE_IMP(hids::vmr9::Renderer, MixerBitmap)
CSP_PROP_2_IMP(FilterConfig, NumberOfStreams, DWORD)
CSP_PROP_2_IMP(FilterConfig, RenderingPrefs, DWORD)

Result FilterConfig::SetRenderingMode(VMR9Mode mode)
{
	CheckInitialized(hids::vmr9::VMRFilterConfig, SetRenderingMode);
	return obj->SetRenderingMode((DWORD)mode);
}

Result FilterConfig::GetRenderingMode(VMR9Mode &mode)
{
	CheckInitialized(hids::vmr9::VMRFilterConfig, GetRenderingMode);
	DWORD m = mode;
	return obj->GetRenderingMode(&m);
}

CSP_PROP_2_IMP(WindowlessControl, BorderColor, COLORREF)
CSP_PROP_2_IMP(WindowlessControl, AspectRatioMode, DWORD)

Result WindowlessControl::GetNativeVideoSize(LONG * lpWidth, LONG * lpHeight, LONG * lpARWidth, LONG * lpARHeight)
{
	CheckInitialized(hids::vmr9::WindowlessControl, GetNativeVideoSize);
	return obj->GetNativeVideoSize(lpWidth, lpHeight, lpARWidth, lpARHeight);
}

Result WindowlessControl::GetMinIdealVideoSize(LONG * lpWidth, LONG * lpHeight)
{
	CheckInitialized(hids::vmr9::WindowlessControl, GetMinIdealVideoSize);
	return obj->GetMinIdealVideoSize(lpWidth, lpHeight);
}

Result WindowlessControl::GetMaxIdealVideoSize(LONG * lpWidth, LONG * lpHeight)
{
	CheckInitialized(hids::vmr9::WindowlessControl, GetMaxIdealVideoSize);
	return obj->GetMaxIdealVideoSize(lpWidth, lpHeight);
}

Result WindowlessControl::SetVideoPosition(RECT &lpSRCRect, RECT &lpDSTRect)
{
	CheckInitialized(hids::vmr9::WindowlessControl, SetVideoPosition);
	return obj->SetVideoPosition(&lpSRCRect, &lpDSTRect);
}

Result WindowlessControl::GetVideoPosition(RECT & lpSRCRect, RECT & lpDSTRect)
{
	CheckInitialized(hids::vmr9::WindowlessControl, GetVideoPosition);
	return obj->GetVideoPosition(&lpSRCRect, &lpDSTRect);
}

Result WindowlessControl::SetVideoClippingWindow(HWND hwnd)
{
	CheckInitialized(hids::vmr9::WindowlessControl, SetVideoClippingWindow);
	CheckArguments(!IsWindow(hwnd), WindowlessControl, SetVideoClippingWindow);
	return obj->SetVideoClippingWindow(hwnd);
}

Result WindowlessControl::RepaintVideo(HWND hwnd, HDC hdc)
{
	CheckInitialized(hids::vmr9::WindowlessControl, RepaintVideo);
	return obj->RepaintVideo(hwnd, hdc);
}

Result WindowlessControl::DisplayModeChanged(void)
{
	CheckInitialized(hids::vmr9::WindowlessControl, DisplayModeChanged);
	return obj->DisplayModeChanged();
}

Result WindowlessControl::GetCurrentImage(BitmapDataHolder &Dib)
{
	CheckInitialized(hids::vmr9::WindowlessControl, GetCurrentImage);
	if (Dib.bmp)
	{
		CoTaskMemFree(Dib.bmp);
		Dib.bmp = nullptr;
	}
	return obj->GetCurrentImage((BYTE**)&Dib.bmp);
}

BitmapDataHolder::BitmapDataHolder()
{
	bmp = nullptr;
}

BitmapDataHolder::~BitmapDataHolder()
{
	if (bmp) CoTaskMemFree(bmp);
}

BITMAPINFOHEADER* BitmapDataHolder::GetBitmap() const
{
	return bmp;
}

Result MixerBitmap::SetAlphaBitmap(const VMR9AlphaBitmap &pBmpParms)
{
	CheckInitialized(hids::vmr9::MixerBitmap, SetAlphaBitmap);
	CheckArguments((pBmpParms.dwFlags & VMR9AlphaBitmap_hDC) != 0 && !pBmpParms.hdc, MixerBitmap, SetAlphaBitmap);
	CheckArguments((pBmpParms.dwFlags & VMR9AlphaBitmap_hDC) == 0 && !pBmpParms.pDDS, MixerBitmap, SetAlphaBitmap);

	HRESULT hr = obj->SetAlphaBitmap(&pBmpParms);
	if (hr == VFW_E_WRONG_STATE)
	{
		DebugInt3;
		OutputDebugString(_T("you must call GetFilterConfig().SetNumberOfStreams prior to call SetAlphaBitmap"));
	}
	return hr;
}

Result MixerBitmap::UpdateAlphaBitmapParameters(const VMR9AlphaBitmap &pBmpParms)
{
	CheckInitialized(hids::vmr9::MixerBitmap, UpdateAlphaBitmapParameters);
	return obj->UpdateAlphaBitmapParameters(&pBmpParms);
}

Result MixerBitmap::GetAlphaBitmapParameters(VMR9AlphaBitmap * pBmpParms)
{
	CheckInitialized(hids::vmr9::MixerBitmap, GetAlphaBitmapParameters);
	return obj->GetAlphaBitmapParameters(pBmpParms);
}
