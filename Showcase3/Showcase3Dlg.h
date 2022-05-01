
// Showcase3Dlg.h : header file
//

#pragma once

#include <mmreg.h>
#include "DShowUtil.h"
#include "smartptr.h"   // smart pointer class
#include "../../../hids/hids.h"
#include "../../../hids/hids_vmr.h"

// CShowcase3Dlg dialog
class CShowcase3Dlg : public CDialogEx
{
// Construction
public:
	CShowcase3Dlg(CWnd* pParent = NULL);	// standard constructor

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_SHOWCASE3_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support

	hids::vmr9::WindowlessControl wc;
	hids::vmr9::MixerBitmap MixerBitmap;
	hids::CaptureGraphBuilder2 pBuilder;
	hids::BaseFilter pVCap;
	hids::GraphBuilder pFg;
	BOOL fPreviewGraphBuilt;
	BOOL fPreviewing;
	hids::Moniker pmVideo;
	WCHAR wachFriendlyName[120];
	HDC crownDC;
	HGDIOBJ crownOldBmp;
	int crownXpos, crownYpos, crown_org_xpos, crown_org_ypos, blink_count, last_instruction2;

	/*
	0: no capturado
	1: capturado para mover
	2: capturado para resizing
	*/
	int crown_mouse_captured_state;
	double crown_module_base;
	float scale_factor;

	bool IsCursorOverCrown(CPoint &point);
	void TakePicture();
	void Init();
	void AddDevices();
	void ChooseDevices(hids::Moniker &pmVideo);
	BOOL StopPreview();
	BOOL StartPreview();
	void TearDownGraph();
	void RemoveDownstream(hids::BaseFilter &pf);
	BOOL BuildPreviewGraph();
	BOOL InitCapFilters();
	void LoadCrownBitmap();
	HWND ghwndApp;
	HDC paintHDC;
	HBITMAP paintOrgBmp;
	RECT paintRect;
	bool isShowingCrown;

// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnDestroy();
	afx_msg void OnBnClickedButton1();
	afx_msg void OnBnClickedHold();
	afx_msg void OnMouseMove(UINT nFlags, CPoint point);
	afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnLButtonUp(UINT nFlags, CPoint point);
	afx_msg void OnTimer(UINT_PTR nIDEvent);
};
