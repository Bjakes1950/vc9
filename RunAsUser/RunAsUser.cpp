// RunAsUser.cpp : Defines the entry point for the application.
//

#include "stdafx.h"
#include "runasuser.h"
#include "resource.h"

typedef struct tagRUNASSTRUCT {
    RUNASTYPE runasType;
    TCHAR szEXE[MAX_PATH+1];
    TCHAR szCmdLine[MAX_PATH+1];
    TCHAR szDomainName[MAX_PATH+1];
    TCHAR szUserName[MAX_PATH+1];
    TCHAR szPassword[MAX_PATH+1];
    TCHAR szDesktop[MAX_PATH+1];
    TCHAR szStartIn[MAX_PATH+1];
    UINT nCmdShow;
    DWORD dwSession;
    BOOL fImpersonate;
    BOOL fElevate;
    HWND hwndMain;
} RUNASSTRUCT, FAR *LPRUNASSTRUCT;

void EditControlAppend(HWND hwndDlg, UINT nId, LPCTSTR szText)
{
   HWND hEdit = GetDlgItem(hwndDlg, nId);
   int ndx = GetWindowTextLength(hEdit);
   SendMessage(hEdit, EM_SETSEL, (WPARAM)ndx, (LPARAM)ndx);
   SendMessage (hEdit, EM_REPLACESEL, 0, (LPARAM) ((LPCTSTR) szText));
}

VOID CALLBACK RunAsUserState(RUNASSTATE state, DWORD dwError, LPCTSTR pszErrorText, LPARAM lParam)
{
    HWND hwndMain = (HWND)lParam;

    TCHAR szText[256] = TEXT("");
    wsprintf(szText, TEXT("RunAsUserState = %d (%d)"), state, dwError);
    EditControlAppend(hwndMain, IDC_RUNASUSER_EXELOG, szText);
    EditControlAppend(hwndMain, IDC_RUNASUSER_EXELOG, TEXT("\r\n"));

    FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, dwError, 0, szText, sizeof(szText)/sizeof(szText[0]), 0);
    EditControlAppend(hwndMain, IDC_RUNASUSER_EXELOG, szText);
    EditControlAppend(hwndMain, IDC_RUNASUSER_EXELOG, TEXT("\r\n"));

    EditControlAppend(hwndMain, IDC_RUNASUSER_EXELOG, pszErrorText);
    EditControlAppend(hwndMain, IDC_RUNASUSER_EXELOG, TEXT("\r\n"));
}

VOID CALLBACK RunAsUserWait(HANDLE hProcess, DWORD dwProcessId, LPARAM lParam)
{
    MSG msg;
    while (WaitForSingleObject(hProcess, 100)==WAIT_TIMEOUT)
    {
        while (PeekMessage(&msg, NULL, 0, 0, PM_REMOVE))
        {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }
}

INT_PTR CALLBACK RunDialogProc(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    switch (uMsg)
    {
        case WM_INITDIALOG:
        {
            LPRUNASSTRUCT lprs = (LPRUNASSTRUCT)lParam;
            SetWindowLongPtr(hwndDlg, GWLP_USERDATA, (LONG_PTR)lprs);
            SetTimer(hwndDlg, 1, 100, NULL);
            return TRUE;
        }

        case WM_TIMER:
        {
            KillTimer(hwndDlg, 1);

            LPRUNASSTRUCT lprs = (LPRUNASSTRUCT)GetWindowLongPtr(hwndDlg, GWLP_USERDATA);
            if (lprs)
            {
                if (RunAsUser_CreateProcess(
                    lprs->runasType, lprs->szEXE, lprs->szCmdLine,
                    lprs->szDomainName, lprs->szUserName, lprs->szPassword,
                    lprs->szDesktop, lprs->szStartIn, lprs->nCmdShow,
                    lprs->dwSession, lprs->fImpersonate, lprs->fElevate,
                    RunAsUserState, RunAsUserWait, (LPARAM)lprs->hwndMain) == FALSE)
                {
                    DWORD dwError = GetLastError();
                    TCHAR szText[256] = TEXT("");
                    wsprintf(szText, TEXT("RunAsUser fail (%d)"), dwError);
                    EditControlAppend(lprs->hwndMain, IDC_RUNASUSER_EXELOG, szText);
                    EditControlAppend(lprs->hwndMain, IDC_RUNASUSER_EXELOG, TEXT("\r\n"));

                    FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, dwError, 0, szText, sizeof(szText)/sizeof(szText[0]), 0);
                    EditControlAppend(lprs->hwndMain, IDC_RUNASUSER_EXELOG, szText);
                    EditControlAppend(lprs->hwndMain, IDC_RUNASUSER_EXELOG, TEXT("\r\n"));
                }
                else
                {
                    EditControlAppend(lprs->hwndMain, IDC_RUNASUSER_EXELOG, TEXT("RunAsUser success."));
                    EditControlAppend(lprs->hwndMain, IDC_RUNASUSER_EXELOG, TEXT("\r\n"));
                }
                EndDialog(hwndDlg, IDOK);
            }
            return TRUE;
        }
    }
    return FALSE;
}

INT_PTR CALLBACK DialogProc(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    switch (uMsg)
    {
        case WM_INITDIALOG:
        {
            SendDlgItemMessage(hwndDlg, IDC_RUNASUSER_TYPE, CB_ADDSTRING, 0, (LPARAM)TEXT("CURRENTUSER"));
            SendDlgItemMessage(hwndDlg, IDC_RUNASUSER_TYPE, CB_ADDSTRING, 0, (LPARAM)TEXT("THISPROCESSUSER"));
            SendDlgItemMessage(hwndDlg, IDC_RUNASUSER_TYPE, CB_ADDSTRING, 0, (LPARAM)TEXT("LOGONLOCALLYUSER")); 
            SendDlgItemMessage(hwndDlg, IDC_RUNASUSER_TYPE, CB_ADDSTRING, 0, (LPARAM)TEXT("PASSWORDLESSUSER"));
            SendDlgItemMessage(hwndDlg, IDC_RUNASUSER_TYPE, CB_ADDSTRING, 0, (LPARAM)TEXT("LOCALSYSTEM"));
            SendDlgItemMessage(hwndDlg, IDC_RUNASUSER_TYPE, CB_SETCURSEL, 0, 0);

            int idx;
            idx = SendDlgItemMessage(hwndDlg, IDC_RUNASUSER_STATE, CB_ADDSTRING, 0, (LPARAM)TEXT("SW_SHOWNORMAL"));
            SendDlgItemMessage(hwndDlg, IDC_RUNASUSER_STATE, CB_SETITEMDATA, idx, (LPARAM)SW_SHOWNORMAL);
            idx = SendDlgItemMessage(hwndDlg, IDC_RUNASUSER_STATE, CB_ADDSTRING, 0, (LPARAM)TEXT("SW_HIDE"));
            SendDlgItemMessage(hwndDlg, IDC_RUNASUSER_STATE, CB_SETITEMDATA, idx, (LPARAM)SW_HIDE);
            idx = SendDlgItemMessage(hwndDlg, IDC_RUNASUSER_STATE, CB_ADDSTRING, 0, (LPARAM)TEXT("SW_SHOWMINIMIZED"));
            SendDlgItemMessage(hwndDlg, IDC_RUNASUSER_STATE, CB_SETITEMDATA, idx, (LPARAM)SW_SHOWMINIMIZED);
            idx = SendDlgItemMessage(hwndDlg, IDC_RUNASUSER_STATE, CB_ADDSTRING, 0, (LPARAM)TEXT("SW_SHOWMAXIMIZED"));
            SendDlgItemMessage(hwndDlg, IDC_RUNASUSER_STATE, CB_SETITEMDATA, idx, (LPARAM)SW_SHOWMAXIMIZED);
            idx = SendDlgItemMessage(hwndDlg, IDC_RUNASUSER_STATE, CB_ADDSTRING, 0, (LPARAM)TEXT("SW_SHOWNOACTIVATE"));
            SendDlgItemMessage(hwndDlg, IDC_RUNASUSER_STATE, CB_SETITEMDATA, idx, (LPARAM)SW_SHOWNOACTIVATE);
            idx = SendDlgItemMessage(hwndDlg, IDC_RUNASUSER_STATE, CB_ADDSTRING, 0, (LPARAM)TEXT("SW_SHOWMINNOACTIVE"));
            SendDlgItemMessage(hwndDlg, IDC_RUNASUSER_STATE, CB_SETITEMDATA, idx, (LPARAM)SW_SHOWMINNOACTIVE);
            idx = SendDlgItemMessage(hwndDlg, IDC_RUNASUSER_STATE, CB_ADDSTRING, 0, (LPARAM)TEXT("SW_SHOWDEFAULT"));
            SendDlgItemMessage(hwndDlg, IDC_RUNASUSER_STATE, CB_SETITEMDATA, idx, (LPARAM)SW_SHOWDEFAULT);
            SendDlgItemMessage(hwndDlg, IDC_RUNASUSER_STATE, CB_SETCURSEL, 0, 0);

            TCHAR szText[256];
            for(int i=0;i<10;i++)
            {
                wsprintf(szText, TEXT("%d"), i);
                SendDlgItemMessage(hwndDlg, IDC_RUNASUSER_SESSION, CB_ADDSTRING, 0, (LPARAM)szText);
            }
            wsprintf(szText, TEXT("%d"), WTSGetActiveConsoleSessionId());
            SendDlgItemMessage(hwndDlg, IDC_RUNASUSER_SESSION, CB_SETCURSEL,
                SendDlgItemMessage(hwndDlg, IDC_RUNASUSER_SESSION, CB_FINDSTRING, -1, (LPARAM)szText), 0);

            SetDlgItemText(hwndDlg, IDC_RUNASUSER_DESKTOP, TEXT("WinSta0\\Default"));

            return TRUE;
        }
        case WM_COMMAND:
        {
            switch (GET_WM_COMMAND_ID(wParam, lParam))
            {
                case IDOK:
                {
                    RUNASSTRUCT rs;

                    TCHAR szText[256];
                    GetDlgItemText(hwndDlg, IDC_RUNASUSER_TYPE, szText, sizeof(szText)/sizeof(szText[0]));
                    if (lstrcmpi(szText, TEXT("THISPROCESSUSER"))==0)
                    {
                        rs.runasType = RUNASTYPE_THISPROCESSUSER;
                    }
                    else if (lstrcmpi(szText, TEXT("LOGONLOCALLYUSER"))==0)
                    {
                        rs.runasType = RUNASTYPE_LOGONLOCALLYUSER;
                    }
                    else if (lstrcmpi(szText, TEXT("PASSWORDLESSUSER"))==0)
                    {
                        rs.runasType = RUNASTYPE_PASSWORDLESSUSER;
                    }
                    else if (lstrcmpi(szText, TEXT("LOCALSYSTEM"))==0)
                    {
                        rs.runasType = RUNASTYPE_LOCALSYSTEM;
                    }
                    else
                    {
                        rs.runasType = RUNASTYPE_CURRENTUSER;
                    }

                    GetDlgItemText(hwndDlg, IDC_RUNASUSER_EXE, rs.szEXE, sizeof(rs.szEXE)/sizeof(rs.szEXE[0]));
                    GetDlgItemText(hwndDlg, IDC_RUNASUSER_CMD, rs.szCmdLine, sizeof(rs.szCmdLine)/sizeof(rs.szCmdLine[0]));
                    GetDlgItemText(hwndDlg, IDC_RUNASUSER_DOMAIN, rs.szDomainName, sizeof(rs.szDomainName)/sizeof(rs.szDomainName[0]));
                    GetDlgItemText(hwndDlg, IDC_RUNASUSER_USERNAME, rs.szUserName, sizeof(rs.szUserName)/sizeof(rs.szUserName[0]));
                    GetDlgItemText(hwndDlg, IDC_RUNASUSER_PASSWORD, rs.szPassword, sizeof(rs.szPassword)/sizeof(rs.szPassword[0]));
                    GetDlgItemText(hwndDlg, IDC_RUNASUSER_DESKTOP, rs.szDesktop, sizeof(rs.szDesktop)/sizeof(rs.szDesktop[0]));
                    GetDlgItemText(hwndDlg, IDC_RUNASUSER_DIR, rs.szStartIn, sizeof(rs.szStartIn)/sizeof(rs.szStartIn[0]));

                    rs.nCmdShow = SendDlgItemMessage(hwndDlg, IDC_RUNASUSER_STATE, CB_GETITEMDATA,
                        SendDlgItemMessage(hwndDlg, IDC_RUNASUSER_STATE, CB_GETCURSEL, 0, 0), 0);

                    GetDlgItemText(hwndDlg, IDC_RUNASUSER_SESSION, szText, sizeof(szText)/sizeof(szText[0]));
                    rs.dwSession = 0;
                    LPTSTR lpszNum = szText;
                    while(*lpszNum)
                    {
                        rs.dwSession = (rs.dwSession<<3) + (rs.dwSession<<1) + (*lpszNum - '0');
                        lpszNum++;
                    }

                    rs.fImpersonate = IsDlgButtonChecked(hwndDlg, IDC_RUNASUSER_IMPERSONATE);
                    rs.fElevate = IsDlgButtonChecked(hwndDlg, IDC_RUNASUSER_ELEVATE);
                    rs.hwndMain = hwndDlg;

                    SetDlgItemText(hwndDlg, IDC_RUNASUSER_EXELOG, TEXT(""));

                    DialogBoxParam((HINSTANCE)GetWindowLongPtr(hwndDlg, GWLP_HINSTANCE),
                        MAKEINTRESOURCE(IDD_DIALOG_RUN), hwndDlg, RunDialogProc, (LPARAM)&rs);
                    break;
                }

                case IDABORT:
                case IDCANCEL:
                {
                    EndDialog(hwndDlg, GET_WM_COMMAND_CMD(wParam, lParam));
                    break;
                }

                case IDC_RUNASUSER_EXE_BROWSE:
                {
                    TCHAR szFile[MAX_PATH+1];
                    ZeroMemory(szFile, sizeof(szFile));

                    OPENFILENAME ofn;
                    ZeroMemory(&ofn, sizeof(OPENFILENAME));
                    ofn.lStructSize = sizeof(OPENFILENAME);
                    ofn.hwndOwner = hwndDlg;
                    ofn.lpstrFilter = TEXT("Excutable files (*.exe)\0*.exe\0");
                    ofn.nFilterIndex = 1;
                    ofn.lpstrFile = szFile;
                    ofn.nMaxFile = MAX_PATH+1;
                    ofn.lpstrTitle = TEXT("Executable File");
                    ofn.Flags = OFN_EXPLORER | OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST | OFN_HIDEREADONLY;

                    if (GetOpenFileName(&ofn))
                    {
                        SetDlgItemText(hwndDlg, IDC_RUNASUSER_EXE, ofn.lpstrFile);
                    }
                    break;
                }

                case IDC_RUNASUSER_DIR_BROWSE:
                {
                    BROWSEINFO bi;
                    ZeroMemory(&bi, sizeof(BROWSEINFO));
                    bi.hwndOwner = hwndDlg;
                    bi.lpszTitle = TEXT("Browse folder");
                    bi.ulFlags = NULL;

                    LPMALLOC ppMalloc;
                    SHGetMalloc(&ppMalloc);
                    LPITEMIDLIST il = SHBrowseForFolder(&bi);
                    if (il != NULL)
                    {
                        TCHAR szFile[MAX_PATH];
                        if (SHGetPathFromIDList(il, szFile))
                        {
                            SetDlgItemText(hwndDlg, IDC_RUNASUSER_DIR, szFile);
                        }
                        ppMalloc->Free(il);
                    }
                    break;
                }
            }
            return TRUE;
        }
    }
    return FALSE;
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow)
{
    // Initialize RunAsUser.
    RUNASUSERCONFIGURATION runasuser;
    runasuser.pszTempServiceName = TEXT("RunAsUser Service");
    runasuser.pszTempServiceExe = NULL;
    runasuser.pszRegistrationKey = NULL;
    RunAsUser_Initialize(&runasuser);

    // Allow RunAsUser DLL to perform some internal operations.
    DWORD dwExitCode = 0;
    if (RunAsUser_CommandLine(GetCommandLine(), &dwExitCode))
        return dwExitCode;

    DialogBox(hInstance, MAKEINTRESOURCE(IDD_RUNASUSER), NULL, DialogProc);
	return 0;
}
