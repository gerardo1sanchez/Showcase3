
// Showcase3Dlg.cpp : implementation file
//

#include "stdafx.h"
#include "Showcase3.h"
#include "Showcase3Dlg.h"
#include "afxdialogex.h"
#include <math.h>

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

bool csp::com_initialized = false;

// CAboutDlg dialog used for App About
#define LongAlign(num)                       (((num + 3) / 4) * 4)

#define CROWN_WIDTH		214
#define CROWN_HEIGHT		140
#define VIDEO_MARGIN 14

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	CFont f;
// Implementation
protected:
	DECLARE_MESSAGE_MAP()
public:
	virtual BOOL OnInitDialog();
};

CAboutDlg::CAboutDlg() : CDialogEx(IDD_ABOUTBOX)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()

CShowcase3Dlg::CShowcase3Dlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(IDD_SHOWCASE3_DIALOG, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CShowcase3Dlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CShowcase3Dlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_WM_DESTROY()
	ON_WM_TIMER()
	ON_BN_CLICKED(IDC_BUTTON1, &CShowcase3Dlg::OnBnClickedButton1)
	ON_BN_CLICKED(IDC_HOLD, &CShowcase3Dlg::OnBnClickedHold)
	ON_WM_MOUSEMOVE()
	ON_WM_LBUTTONDOWN()
	ON_WM_LBUTTONUP()
	ON_WM_TIMER()
END_MESSAGE_MAP()

BOOL CShowcase3Dlg::StopPreview() {
	if (!fPreviewing) return FALSE;

	if (!pFg.GetMediaControl().Stop()) return false;

	fPreviewing = FALSE;
	::InvalidateRect(ghwndApp, nullptr, TRUE);

	return TRUE;
}

BOOL CShowcase3Dlg::BuildPreviewGraph()
{
	// we have one already
	if (fPreviewGraphBuilt)
		return TRUE;

	// No rebuilding while we're running
	if (fPreviewing)
		return FALSE;

	if (!pVCap.IsValid()) return FALSE;

	hids::vmr9::Renderer vmrRenderer;
	vmrRenderer.CreateInstance();
	auto fc = vmrRenderer.GetFilterConfig();
	fc.NumberOfStreams = 1;
	fc.SetRenderingMode(VMR9Mode_Windowless);
	wc = vmrRenderer.GetWindowlessControl();
	wc.SetVideoClippingWindow(ghwndApp);

	pFg.AddFilter(vmrRenderer, L"Video Mixer Renderer");

	if(!pBuilder.RenderStream(&PIN_CATEGORY_CAPTURE, &MEDIATYPE_Video, pVCap, nullptr, &vmrRenderer)) return false;

	hids::AMStreamConfig pVSC;
	pVSC = pBuilder.FindInterface(&PIN_CATEGORY_CAPTURE, &MEDIATYPE_Interleaved, pVCap, IID_IAMStreamConfig).Cast<hids::AMStreamConfig>();
	if (!pVSC) pVSC = pBuilder.FindInterface(&PIN_CATEGORY_CAPTURE, &MEDIATYPE_Video, pVCap, IID_IAMStreamConfig).Cast<hids::AMStreamConfig>();

	hids::AmMediaTypeHolder media;
	pVSC.GetFormat(media);
	VIDEOINFOHEADER* vih = media.GetVideoInfoHeader();

	RECT s, d, rc;
	::GetClientRect(ghwndApp, &rc);
	rc.bottom -= VIDEO_MARGIN;
	float factor = (float)rc.bottom / (float)vih->bmiHeader.biHeight;
	rc.right = (long)((float)vih->bmiHeader.biWidth * factor);
	SetRect(&s, 0, 0, vih->bmiHeader.biWidth, vih->bmiHeader.biHeight);
	SetRect(&d, VIDEO_MARGIN, VIDEO_MARGIN, rc.right, rc.bottom);
	wc.SetVideoPosition(s, d);

	VMR9AlphaBitmap x;
	memset(&x, 0, sizeof(x));
	LoadCrownBitmap();
	x.hdc = crownDC;
	SetRect(&x.rSrc, 0, 0, CROWN_WIDTH, CROWN_HEIGHT);
	paintRect.right = rc.right;
	paintRect.bottom = rc.bottom;
	crownXpos = 0;
	crownYpos = 0;
	scale_factor = 1.0f;
	x.rDest.left = (float)crownXpos / (float)paintRect.right;
	x.rDest.top = (float)crownYpos / (float)paintRect.bottom;
	x.rDest.right = x.rDest.left + (float)(CROWN_WIDTH) * scale_factor  / (float)paintRect.right;
	x.rDest.bottom = x.rDest.top + (float)(CROWN_HEIGHT)* scale_factor / (float)paintRect.bottom;
	x.fAlpha = 1.0f;
	x.dwFlags = VMR9AlphaBitmap_hDC | VMR9AlphaBitmap_SrcColorKey;
	isShowingCrown = true;

	MixerBitmap = vmrRenderer.GetMixerBitmap();
	MixerBitmap.SetAlphaBitmap(x);

	fPreviewGraphBuilt = TRUE;
	return TRUE;
}

void CShowcase3Dlg::LoadCrownBitmap()
{
	if (crownDC) return;
	HANDLE b = LoadImage(GetModuleHandle(nullptr), MAKEINTRESOURCE(IDB_CROWN), IMAGE_BITMAP, CROWN_WIDTH * 2, CROWN_HEIGHT, LR_CREATEDIBSECTION);
	HDC aux = ::GetDC(ghwndApp);
	crownDC = CreateCompatibleDC(aux);
	::ReleaseDC(ghwndApp, aux);
	crownOldBmp = SelectObject(crownDC, b);
}

BOOL CShowcase3Dlg::InitCapFilters()
{
	if (pmVideo.IsValid()) {
		hids::PropertyBag pBag;
		wachFriendlyName[0] = 0;

		pBag = pmVideo.BindToStoragePropertyBag();
		if (pBag) {
			VARIANT var;
			var.vt = VT_BSTR;

			if (pBag.Read(L"FriendlyName", var, nullptr).Succeeded())
			{
				StringCchCopyW(wachFriendlyName, sizeof(wachFriendlyName) / sizeof(wachFriendlyName[0]), var.bstrVal);
				SysFreeString(var.bstrVal);
			}
		}

		pVCap = pmVideo.BindToObjectBaseFilter();
		if (!pVCap) return false;
	}

	if (!pVCap) return false;

	if (!pFg.CreateInstance()) return false;

	if (!pBuilder.SetFiltergraph(pFg)) return false;

	if (!pFg.AddFilter(pVCap, wachFriendlyName)) return false;

	BuildPreviewGraph();

	return TRUE;
}

#define VERSIZE 40
#define DESCSIZE 80

void CShowcase3Dlg::ChooseDevices(hids::Moniker &video) {
	int versize = VERSIZE;
	int descsize = DESCSIZE;
	WCHAR wachVer[VERSIZE] = { 0 }, wachDesc[DESCSIZE] = { 0 };
	TCHAR tachStatus[VERSIZE + DESCSIZE + 5] = { 0 };


	// they chose a new device. rebuild the graphs
	if (pmVideo != video) {

		//pmVideo.Release();
		pmVideo = video;

		if (fPreviewing)
			StopPreview();
		if (fPreviewGraphBuilt)
			TearDownGraph();

		InitCapFilters();
		BuildPreviewGraph();
	}
}

void CShowcase3Dlg::AddDevices()
{
	hids::SystemDeviceEnum sde;
	
	for (auto a : sde.CreateClassEnumerator(CLSID_VideoInputDeviceCategory, 0))
	{
		if (!a.BindToStoragePropertyBag().IsValid()) continue;
		ChooseDevices(a);
		return;
	}
}

void CShowcase3Dlg::Init()
{
	ghwndApp = /*GetDlgItem(IDC_CAPTURE_CTRL)->*/GetSafeHwnd();
	CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
	fPreviewGraphBuilt = fPreviewing = FALSE;
	AddDevices();
}

void CShowcase3Dlg::RemoveDownstream(hids::BaseFilter &pf)
{
	if (!pf) return;

	hids::Pin pTo;
	hids::PinInfo pininfo;

	for (auto pin : pf.Pins())
	{
		pTo = pin.ConnectedTo();
		if (!pTo)
		{
			pin.QueryPinInfo(pininfo);
			continue;
		}
		if (!pTo.QueryPinInfo(pininfo)) continue;
		if (pininfo.dir != PINDIR_INPUT) continue;
		RemoveDownstream(pininfo.pFilter);
		pFg.Disconnect(pTo);
		pFg.Disconnect(pin);
		pFg.RemoveFilter(pininfo.pFilter);
	}
}

void CShowcase3Dlg::TearDownGraph()
{
	if (pVCap) RemoveDownstream(pVCap);
	fPreviewGraphBuilt = FALSE;
}

BOOL CShowcase3Dlg::StartPreview() {
	if (fPreviewing) return TRUE;
	if (!fPreviewGraphBuilt) return FALSE;
	auto ret = pFg.GetMediaControl().Run();
	fPreviewing = ret.Succeeded() || ret.last_error == S_FALSE/* acording to SDK, S_FALSE isn't a error, it means graph is still starting */;
	return fPreviewing;
}

void CShowcase3Dlg::TakePicture()
{
	BITMAPFILEHEADER fh;
	HANDLE file;
	DWORD written;
	TCHAR dest_path[MAX_PATH];
	OPENFILENAME ofn;
	hids::vmr9::BitmapDataHolder bmp;

	wc.GetCurrentImage(bmp);

	dest_path[0] = 0;
	memset(&ofn, 0, sizeof(OPENFILENAME));
	ofn.lStructSize = sizeof(OPENFILENAME);
	ofn.hwndOwner = *this;
	ofn.lpstrFilter = L"Bitmap\0*.bmp\0\0";
	ofn.lpstrFile = dest_path;
	ofn.nMaxFile = MAX_PATH;
	ofn.lpstrTitle = L"Save Bitmap";
	ofn.Flags = OFN_OVERWRITEPROMPT | OFN_PATHMUSTEXIST;
	ofn.lpstrDefExt = L"bmp";
	if (!GetSaveFileName(&ofn)) return;

	file = CreateFile(dest_path, GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
	fh.bfOffBits = sizeof(BITMAPFILEHEADER) + sizeof(BITMAPINFOHEADER);
	fh.bfReserved1 = fh.bfReserved2 = 0;
	fh.bfSize = sizeof(BITMAPFILEHEADER);
	fh.bfType = 0x4D42;
	BITMAPINFOHEADER *bih = bmp.GetBitmap();
	WriteFile(file, &fh, sizeof(BITMAPFILEHEADER), &written, NULL);
	WriteFile(file, bih, sizeof(BITMAPINFOHEADER) + sizeof(RGBQUAD) * bih->biClrUsed + LongAlign((bih->biBitCount / 8) * bih->biWidth) * bih->biHeight, &written, nullptr);
	CloseHandle(file);
}

BOOL CShowcase3Dlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// Add "About..." menu item to system menu.

	// IDM_ABOUTBOX must be in the system command range.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

	// TODO: Add extra initialization here
#if !defined(_DEBUG)
	CAboutDlg dlgAbout;
	dlgAbout.DoModal();
#endif

	crownDC = nullptr;
	isShowingCrown = false;
	Init();
	StartPreview();
	blink_count = last_instruction2 = 0;
	SetTimer(0, 500, nullptr);
	SetDlgItemText(IDC_EDIT2, L"Instructions:\r\nTry to move crown using mouse's left button");

	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CShowcase3Dlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CShowcase3Dlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else 
	{
		CDialogEx::OnPaint();
	}
}

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CShowcase3Dlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

void CShowcase3Dlg::OnDestroy()
{
	StopPreview();
	TearDownGraph();

	if (crownDC)
	{
		DeleteObject(SelectObject(crownDC, crownOldBmp));
		DeleteDC(crownDC);
	}
	CDialogEx::OnDestroy();
}

BOOL CAboutDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  Add extra initialization here

	SetDlgItemText(IDC_EDIT1, L"This is a free software, THE SOFTWARE IS PROVIDED 'AS IS', WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.");
	SetDlgItemText(IDC_EDIT30, L"http://portfolio.ambar.com.ve\t\tgerardo1sanchez@gmail.com\r\n\r\nThis software is a DEMO, intended to demostrates MY SKILLS in this subjects:\r\n\r\n* DirectShow Video Basics\t* MFC dialog\r\n* Image Manipulation\t* Bitmap File Creation");
	f.CreateFont(20, 0, 0, 0, FW_BOLD, 0, 0, 0, 0, 0, 0, 0, 0, L"Verdana");
	SendDlgItemMessage(IDC_LABEL51, WM_SETFONT, (WPARAM)(HFONT)f, true);

	return TRUE;  // return TRUE unless you set the focus to a control
					  // EXCEPTION: OCX Property Pages should return FALSE
}

void CShowcase3Dlg::OnBnClickedButton1()
{
	TakePicture();
}

void CShowcase3Dlg::OnBnClickedHold()
{
	if (fPreviewing)
	{
		if (pFg.GetMediaControl().StopWhenReady().Succeeded())
		{
			SetDlgItemText(IDC_HOLD, L"Restart Camera");
			fPreviewing = false;
		}
	}
	else
	{
		auto ret = pFg.GetMediaControl().Run();
		if (ret.Succeeded() || ret.last_error == S_FALSE/* acording to SDK, S_FALSE isn't a error, it means graph is still starting */) {
			fPreviewing = true;
			SetDlgItemText(IDC_HOLD, L"Hold Image");
		}
	}
}

void CShowcase3Dlg::OnMouseMove(UINT nFlags, CPoint point)
{
	CDialogEx::OnMouseMove(nFlags, point);
	double m;
	switch (crown_mouse_captured_state)
	{
	case 0:
		if (IsCursorOverCrown(point)) SetCursor(LoadCursor(nullptr, (nFlags & MK_CONTROL)?IDC_SIZEALL:IDC_HAND));
		return;
	case 1:
		crownXpos = crown_org_xpos + point.x;
		crownYpos = crown_org_ypos + point.y;
		SetCursor(LoadCursor(nullptr, IDC_HAND));
		break;
	case 2:
		m = sqrt(pow((double)(crownXpos - point.x), 2) + pow((double)(crownYpos - point.y), 2));
		scale_factor = (float)(m / crown_module_base);
		if (scale_factor < 0.2f)
		{
			scale_factor = 0.2f;
		}
		else if (scale_factor > 3.0f)
		{
			scale_factor = 3.0f;
		}
		SetCursor(LoadCursor(nullptr, IDC_SIZEALL));
	}

	VMR9AlphaBitmap x;

	memset(&x, 0, sizeof(x));
	x.rDest.left = (float)crownXpos / (float)paintRect.right;
	x.rDest.top = (float)crownYpos / (float)paintRect.bottom;
	x.rDest.right = x.rDest.left + (float)(CROWN_WIDTH)* scale_factor / (float)paintRect.right;
	x.rDest.bottom = x.rDest.top + (float)(CROWN_HEIGHT)* scale_factor / (float)paintRect.bottom;
	x.fAlpha = 1.0f;
	x.dwFlags = VMR9AlphaBitmap_SrcColorKey;
	MixerBitmap.UpdateAlphaBitmapParameters(x);
}

bool CShowcase3Dlg::IsCursorOverCrown(CPoint &point)
{
	if (!isShowingCrown || !fPreviewing) return false;
	if (point.x - VIDEO_MARGIN < crownXpos) return false;
	if (point.y - VIDEO_MARGIN < crownYpos) return false;
	if (point.x - VIDEO_MARGIN >= crownXpos + (int)((float)CROWN_WIDTH * scale_factor)) return false;
	if (point.y - VIDEO_MARGIN >= crownYpos + (int)((float)CROWN_HEIGHT * scale_factor)) return false;
	return 0 == GetPixel(crownDC, (int)((float)(point.x - VIDEO_MARGIN - crownXpos) / scale_factor) + CROWN_WIDTH, (int)((float)(point.y - VIDEO_MARGIN - crownYpos) / scale_factor));
}

void CShowcase3Dlg::OnLButtonDown(UINT nFlags, CPoint point)
{
	CDialogEx::OnLButtonDown(nFlags, point);
	if (!IsCursorOverCrown(point)) return;
	crown_org_xpos = crownXpos - point.x;
	crown_org_ypos = crownYpos - point.y;
	crown_mouse_captured_state = (nFlags & MK_CONTROL)?2:1;
	if (crown_mouse_captured_state == 2) crown_module_base = sqrt(pow((double)(crown_org_xpos), 2) + pow((double)(crown_org_ypos), 2)) / (double)scale_factor;
	SetCapture();
}

void CShowcase3Dlg::OnLButtonUp(UINT nFlags, CPoint point)
{
	switch (crown_mouse_captured_state)
	{
	case 1:
		if (last_instruction2 == 0)
		{
			last_instruction2 = 1;
			SetDlgItemText(IDC_EDIT2, L"Instructions:\r\nTry to resize crown pressing AND HOLD keyboard CTRL button before pressing mouse's left button");
		}
		break;
	case 2:
		if (last_instruction2 == 1)
		{
			last_instruction2 = 2;
			SetDlgItemText(IDC_EDIT2, L"Instructions:\r\nPress 'Save picture to file' to Save your picture to bitmap file\r\n\r\nSuggestion: press 'Hold Image' button to capture a picture before save it");
		}
	}
	crown_mouse_captured_state = 0;
	ReleaseCapture();
	CDialogEx::OnLButtonUp(nFlags, point);
}


void CShowcase3Dlg::OnTimer(UINT_PTR nIDEvent)
{
	GetDlgItem(IDC_EDIT2)->ShowWindow(((blink_count++) % 2) ? SW_SHOW : SW_HIDE);
	if (blink_count > ((last_instruction2 == 0)?5:3)) KillTimer(nIDEvent);

	CDialogEx::OnTimer(nIDEvent);
}
