
#include <d3d9.h>
#include <vmr9.h>

using namespace csp;

namespace hids
{
	namespace vmr9
	{
		class BitmapDataHolder
		{
		public:
			BitmapDataHolder();
			~BitmapDataHolder();

			BITMAPINFOHEADER* GetBitmap() const;

		private:
			BitmapDataHolder(const BitmapDataHolder&) {}
			void operator=(const BitmapDataHolder&) {}
			BITMAPINFOHEADER* bmp;
			friend class WindowlessControl;
		};

		class WindowlessControl: public Base<IVMRWindowlessControl9>
		{
		public:
			CSP_PROP_2(BorderColor, COLORREF)
			CSP_PROP_2(AspectRatioMode, DWORD)
			Result GetNativeVideoSize(LONG *lpWidth, LONG *lpHeight, LONG *lpARWidth, LONG *lpARHeight);
			Result GetMinIdealVideoSize(LONG *lpWidth, LONG *lpHeight);
			Result GetMaxIdealVideoSize(LONG *lpWidth, LONG *lpHeight);
			Result SetVideoPosition(RECT &lpSRCRect, RECT &lpDSTRect);
			Result GetVideoPosition(RECT &lpSRCRect, RECT &lpDSTRect);
			Result SetVideoClippingWindow(HWND hwnd);
			Result RepaintVideo(HWND hwnd, HDC hdc);
			Result DisplayModeChanged();
			Result GetCurrentImage(BitmapDataHolder &Dib);

			friend class Renderer;
			GENERIC_DEFINITIONS(WindowlessControl)
		};
									
		class FilterConfig : public Base<IVMRFilterConfig9>
		{
		public:
			CSP_PROP_2(NumberOfStreams, DWORD)
			CSP_PROP_2(RenderingPrefs, DWORD)
			CSP_PROP_2(RenderingMode, VMR9Mode)

			friend class Renderer;
			GENERIC_DEFINITIONS(FilterConfig)
		};

		class MixerBitmap: public Base<IVMRMixerBitmap9>
		{
		public:
			Result SetAlphaBitmap(const VMR9AlphaBitmap &pBmpParms);
			Result UpdateAlphaBitmapParameters(const VMR9AlphaBitmap &pBmpParms);
			Result GetAlphaBitmapParameters(VMR9AlphaBitmap *pBmpParms);
			friend class Renderer;
			GENERIC_DEFINITIONS(MixerBitmap)
		};

		class Renderer : public BaseFilter
		{
		public:
			typedef BaseFilter BASE_CLASS;
			Result CreateInstance();
			CSP_QUERY_INTERFACE(FilterConfig)
			CSP_QUERY_INTERFACE(WindowlessControl)
			CSP_QUERY_INTERFACE(MixerBitmap)
			GENERIC_DEFINITIONS(Renderer)
		};
	}
}