// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
// PARTICULAR PURPOSE.
//
// Copyright (c) Microsoft Corporation. All rights reserved.

//------------------------------------------------------------------------------
// File: PlayCap.cpp
//
// Desc: DirectShow sample code - a very basic application using video capture
//       devices.  It creates a window and uses the first available capture
//       device to render and preview video capture data.
//------------------------------------------------------------------------------

#define _WIN32_WINNT 0x0500

#include <windows.h>
#include <dshow.h>
#include <stdio.h>
#include <strsafe.h>
#include "..\..\..\hids.h"
#include "PlayCap.h"

// An application can advertise the existence of its filter graph
// by registering the graph with a global Running Object Table (ROT).
// The GraphEdit application can detect and remotely view the running
// filter graph, allowing you to 'spy' on the graph with GraphEdit.
//
// To enable registration in this sample, define REGISTER_FILTERGRAPH.
//
#define REGISTER_FILTERGRAPH


//
// Global data
//
HWND ghApp=0;
DWORD g_dwGraphRegister=0;

hids::VideoWindow  g_pVW;
hids::MediaControl g_pMC;
hids::MediaEventEx g_pME;
hids::GraphBuilder g_pGraph;
hids::CaptureGraphBuilder2 g_pCapture;
PLAYSTATE g_psCurrent = Stopped;

bool CaptureVideo()
{
    csp::LastError hr;
	 hids::BaseFilter pSrcFilter;

    // Get DirectShow interfaces
    if (!GetInterfaces())
    {
        Msg(TEXT("Failed to get video interfaces!  hr=0x%x"), hr);
        return false;
    }

    // Attach the filter graph to the capture graph
    if (!g_pCapture.SetFiltergraph(g_pGraph))
    {
        Msg(TEXT("Failed to set capture filter graph!  hr=0x%x"), hr);
        return false;
    }

    // Use the system device enumerator and class enumerator to find
    // a video capture/preview device, such as a desktop USB video camera.
	 if (!FindCaptureDevice(pSrcFilter)) return false;
   
    // Add Capture filter to our graph.
    if (!g_pGraph.AddFilter(pSrcFilter, L"Video Capture"))
    {
        Msg(TEXT("Couldn't add the capture filter to the graph!  hr=0x%x\r\n\r\n") 
            TEXT("If you have a working video capture device, please make sure\r\n")
            TEXT("that it is connected and is not being used by another application.\r\n\r\n")
            TEXT("The sample will now close."), hr);
        return false;
    }

    // Render the preview pin on the video capture filter
    // Use this instead of g_pGraph->RenderFile
    if (!g_pCapture.RenderStream(&PIN_CATEGORY_PREVIEW, &MEDIATYPE_Video, pSrcFilter, nullptr, nullptr))
    {
        Msg(TEXT("Couldn't render the video capture stream.  hr=0x%x\r\n")
            TEXT("The capture device may already be in use by another application.\r\n\r\n")
            TEXT("The sample will now close."), hr);
        return false;
    }

    // Set video window style and position
    if (!SetupVideoWindow())
    {
        Msg(TEXT("Couldn't initialize video window!  hr=0x%x"), hr);
        return false;
    }

#ifdef REGISTER_FILTERGRAPH
    // Add our graph to the running object table, which will allow
    // the GraphEdit application to "spy" on our graph
    if (!AddGraphToRot(g_pGraph, g_dwGraphRegister))
    {
        Msg(TEXT("Failed to register filter graph with ROT!  hr=0x%x"), hr);
        g_dwGraphRegister = 0;
    }
#endif

    // Start previewing video data
    if (!g_pMC.Run())
    {
        Msg(TEXT("Couldn't run the graph!  hr=0x%x"), hr);
        return false;
    }

    // Remember current state
    g_psCurrent = Running;
        
    return true;
}

bool FindCaptureDevice(hids::BaseFilter &ppSrcFilter)
{
    csp::LastError hr;
	 hids::Moniker pMoniker;
	 hids::SystemDeviceEnum pDevEnum;
	 hids::EnumMoniker pClassEnum;
   
    // Create the system device enumerator
    if (!pDevEnum.CreateInstance())
    {
        Msg(TEXT("Couldn't create system enumerator!  hr=0x%x"), hr);
    }

    // Create an enumerator for the video capture devices
	 for (auto v : pDevEnum.CreateClassEnumerator(CLSID_VideoInputDeviceCategory, 0))
	 {
		 ppSrcFilter = v.BindToObjectBaseFilter();
		 if (ppSrcFilter) return true;
	 }

	// If there are no enumerators for the requested type, then 
	// CreateClassEnumerator will succeed, but pClassEnum will be NULL.
	MessageBox(ghApp,TEXT("No video capture device was detected.\r\n\r\n")
		TEXT("This sample requires a video capture device, such as a USB WebCam,\r\n")
		TEXT("to be installed and working properly.  The sample will now close."),
		TEXT("No Video Capture Hardware"), MB_OK | MB_ICONINFORMATION);
	return false;
}


bool GetInterfaces(void)
{
    // Create the filter graph
	if (!g_pGraph.CreateInstance()) return false;

    // Create the capture graph builder
	if (!g_pCapture.CreateInstance()) return false;
    
   // Obtain interfaces for media control and Video Window
	g_pMC = g_pGraph.GetMediaControl();
	if (!g_pMC) return false;

	g_pVW = g_pGraph.GetVideoWindow();
	if (!g_pVW) return false;

	g_pME = g_pGraph.GetMediaEventEx();
	if (!g_pME) return false;

    // Set the window handle used to process graph events
    if(!g_pME.SetNotifyWindow((OAHWND)ghApp, WM_GRAPHNOTIFY, 0)) return false;

    return true;
}


void CloseInterfaces(void)
{
    // Stop previewing data
    if (g_pMC) g_pMC.StopWhenReady();

    g_psCurrent = Stopped;

    // Stop receiving events
    if (g_pME)  g_pME.SetNotifyWindow(NULL, WM_GRAPHNOTIFY, 0);

    // Relinquish ownership (IMPORTANT!) of the video window.
    // Failing to call put_Owner can lead to assert failures within
    // the video renderer, as it still assumes that it has a valid
    // parent window.
    if(g_pVW)
    {
        g_pVW.put_Visible(OAFALSE);
        g_pVW.put_Owner(NULL);
    }

#ifdef REGISTER_FILTERGRAPH
    // Remove filter graph from the running object table   
    if (g_dwGraphRegister)
        RemoveGraphFromRot(g_dwGraphRegister);
#endif

    // Release DirectShow interfaces
	 // this is required because objects were declared globally (its necesary to release them before call CoUninitialize)
    g_pMC.Release();
    g_pME.Release();
    g_pVW.Release();
    g_pGraph.Release();
    g_pCapture.Release();
}


bool SetupVideoWindow(void)
{
    // Set the video window to be a child of the main window
    if(!g_pVW.put_Owner((OAHWND)ghApp)) return false;
    
    // Set video window style
    if(!g_pVW.put_WindowStyle(WS_CHILD | WS_CLIPCHILDREN)) return false;

    // Use helper function to position video window in client rect 
    // of main application window
    ResizeVideoWindow();

    // Make the video window visible, now that it is properly positioned
    if(!g_pVW.put_Visible(OATRUE)) return false;

    return true;
}


void ResizeVideoWindow(void)
{
    // Resize the video preview window to match owner window size
	if (!g_pVW) return;
   RECT rc;
        
   // Make the preview video fill our window
   GetClientRect(ghApp, &rc);
   g_pVW.SetWindowPosition(0, 0, rc.right, rc.bottom);
}


bool ChangePreviewState(int nShow)
{
    // If the media control interface isn't ready, don't call it
    if (!g_pMC) return false;
    
    if (nShow)
    {
        if (g_psCurrent != Running)
        {
            // Start previewing video data
			  if (!g_pMC.Run()) return false;
            g_psCurrent = Running;
        }
    }
    else
    {
        // Stop previewing video data
        if(!g_pMC.StopWhenReady()) return false;
        g_psCurrent = Stopped;
    }

    return true;
}


#ifdef REGISTER_FILTERGRAPH

bool AddGraphToRot(hids::GraphBuilder &pUnkGraph, DWORD &pdwRegister) 
{

    if (!pUnkGraph) return false;

    hids::Moniker pMoniker;
    WCHAR wsz[128];
	 lstrcpyW(wsz, L"hids Capture example");
    if (!pMoniker.CreateItemMoniker(L"!", wsz))
    {
			// Use the ROTFLAGS_REGISTRATIONKEEPSALIVE to ensure a strong reference
			// to the object.  Using this flag will cause the object to remain
			// registered until it is explicitly revoked with the Revoke() method.
			//
			// Not using this flag means that if GraphEdit remotely connects
			// to this graph and then GraphEdit exits, this object registration 
			// will be deleted, causing future attempts by GraphEdit to fail until
			// this application is restarted or until the graph is registered again.
			hids::RunningObjectTable pROT;
			if(!pROT.Register(ROTFLAGS_REGISTRATIONKEEPSALIVE, pUnkGraph.Cast<csp::Unknown>(), pMoniker, pdwRegister)) return false;
    }

    return true;
}


// Removes a filter graph from the Running Object Table
void RemoveGraphFromRot(DWORD pdwRegister)
{
	hids::RunningObjectTable pROT;
	pROT.Revoke(pdwRegister);
}

#endif


void Msg(TCHAR *szFormat, ...)
{
    TCHAR szBuffer[1024];  // Large buffer for long filenames or URLs
    const size_t NUMCHARS = sizeof(szBuffer) / sizeof(szBuffer[0]);
    const int LASTCHAR = NUMCHARS - 1;

    // Format the input string
    va_list pArgs;
    va_start(pArgs, szFormat);

    // Use a bounded buffer size to prevent buffer overruns.  Limit count to
    // character size minus one to allow for a NULL terminating character.
    (void)StringCchVPrintf(szBuffer, NUMCHARS - 1, szFormat, pArgs);
    va_end(pArgs);

    // Ensure that the formatted string is NULL-terminated
    szBuffer[LASTCHAR] = TEXT('\0');

    MessageBox(NULL, szBuffer, TEXT("PlayCap Message"), MB_OK | MB_ICONERROR);
}


void HandleGraphEvent(void)
{
    LONG evCode;
	LONG_PTR evParam1, evParam2;

    if (!g_pME) return;

    while(g_pME.GetEvent(evCode, &evParam1, &evParam2, 0).Succeeded())
    {
        //
        // Free event parameters to prevent memory leaks associated with
        // event parameter data.  While this application is not interested
        // in the received events, applications should always process them.
        //
        g_pME.FreeEventParams(evCode, evParam1, evParam2);
        
        // Insert event processing code here, if desired
    }
}


LRESULT CALLBACK WndMainProc (HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
        case WM_GRAPHNOTIFY:
            HandleGraphEvent();
            break;

        case WM_SIZE:
            ResizeVideoWindow();
            break;

        case WM_WINDOWPOSCHANGED:
            ChangePreviewState(! (IsIconic(hwnd)));
            break;
        case WM_CLOSE:            
            // Hide the main window while the graph is destroyed
            ShowWindow(ghApp, SW_HIDE);
            CloseInterfaces();  // Stop capturing and release interfaces
            break;

        case WM_DESTROY:
            PostQuitMessage(0);
            return 0;
    }

    // Pass this message to the video window for notification of system changes
    if (g_pVW) g_pVW.NotifyOwnerMessage((LONG_PTR) hwnd, message, wParam, lParam);

    return DefWindowProc (hwnd , message, wParam, lParam);
}


int PASCAL WinMain(HINSTANCE hInstance, HINSTANCE hInstP, LPSTR lpCmdLine, int nCmdShow)
{
    MSG msg={0};
    WNDCLASS wc;

    // Initialize COM
    if(FAILED(CoInitializeEx(NULL, COINIT_APARTMENTTHREADED)))
    {
        Msg(TEXT("CoInitialize Failed!\r\n"));   
        exit(1);
    } 

    // Register the window class
    ZeroMemory(&wc, sizeof wc);
    wc.lpfnWndProc   = WndMainProc;
    wc.hInstance     = hInstance;
    wc.lpszClassName = CLASSNAME;
    wc.lpszMenuName  = NULL;
    wc.hbrBackground = (HBRUSH)GetStockObject(BLACK_BRUSH);
    wc.hCursor       = LoadCursor(NULL, IDC_ARROW);
    wc.hIcon         = NULL;
    if(!RegisterClass(&wc))
    {
        Msg(TEXT("RegisterClass Failed! Error=0x%x\r\n"), GetLastError());
        CoUninitialize();
        exit(1);
    }

    // Create the main window.  The WS_CLIPCHILDREN style is required.
    ghApp = CreateWindow(CLASSNAME, APPLICATIONNAME,
                         WS_OVERLAPPEDWINDOW | WS_CAPTION | WS_CLIPCHILDREN,
                         CW_USEDEFAULT, CW_USEDEFAULT,
                         DEFAULT_VIDEO_WIDTH, DEFAULT_VIDEO_HEIGHT,
                         0, 0, hInstance, 0);

    if(ghApp)
    {
        HRESULT hr;

        // Create DirectShow graph and start capturing video
        hr = CaptureVideo();
        if (FAILED (hr))
        {
            CloseInterfaces();
            DestroyWindow(ghApp);
        }
        else
        {
            // Don't display the main window until the DirectShow
            // preview graph has been created.  Once video data is
            // being received and processed, the window will appear
            // and immediately have useful video data to display.
            // Otherwise, it will be black until video data arrives.
            ShowWindow(ghApp, nCmdShow);
        }       

        // Main message loop
        while(GetMessage(&msg,NULL,0,0))
        {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

    // Release COM
    CoUninitialize();

    return (int) msg.wParam;
}



