/**
 * WinPE Image Tool - DISM GUI Frontend
 * 
 * Features: Image Capture, Apply, Info
 * Backend: DISM.exe
 * Target: Windows PE (native Win32 API, no runtime dependencies)
 * 
 * Build: cl /O2 /W4 /EHsc /DUNICODE /D_UNICODE main.cpp /link /SUBSYSTEM:WINDOWS
 *        user32.lib gdi32.lib comctl32.lib comdlg32.lib shell32.lib ole32.lib
 */

#define WIN32_LEAN_AND_MEAN
#ifndef UNICODE
#define UNICODE
#endif
#ifndef _UNICODE
#define _UNICODE
#endif

#include <windows.h>
#include <commctrl.h>
#include <commdlg.h>
#include <shlobj.h>
#include <shellapi.h>
#include <stdio.h>
#include <string>
#include <vector>
#include <thread>
#include <atomic>

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "comdlg32.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(linker, "/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

// ============================================================
// Constants & IDs
// ============================================================
#define APP_TITLE       L"WinPE Image Tool"
#define APP_WIDTH       750
#define APP_HEIGHT      580

// Tab pages
#define TAB_CAPTURE     0
#define TAB_APPLY       1
#define TAB_INFO        2

// Common
#define IDC_TAB             1000
#define IDC_LOG             1001
#define IDC_PROGRESS        1002
#define IDC_STATUS          1003

// Capture tab
#define IDC_CAP_SRCVOL      1100
#define IDC_CAP_REFRESH     1101
#define IDC_CAP_DSTPATH     1102
#define IDC_CAP_DSTBROWSE   1103
#define IDC_CAP_NAME        1104
#define IDC_CAP_DESC        1105
#define IDC_CAP_COMPRESS    1106
#define IDC_CAP_START       1107

// Apply tab
#define IDC_APP_WIMPATH     1200
#define IDC_APP_WIMBROWSE   1201
#define IDC_APP_INDEX       1202
#define IDC_APP_DSTVOL      1203
#define IDC_APP_REFRESH     1204
#define IDC_APP_START       1205

// Info tab
#define IDC_INF_WIMPATH     1300
#define IDC_INF_WIMBROWSE   1301
#define IDC_INF_START       1302
#define IDC_INF_RESULT      1303

// Custom messages
#define WM_APPEND_LOG       (WM_USER + 1)
#define WM_SET_PROGRESS     (WM_USER + 2)
#define WM_OPERATION_DONE   (WM_USER + 3)

// ============================================================
// Globals
// ============================================================
static HINSTANCE g_hInst;
static HWND g_hWnd;
static HWND g_hTab;
static HWND g_hLog;
static HWND g_hProgress;
static HWND g_hStatus;

// Capture controls
static HWND g_hCapSrcVol, g_hCapDstPath, g_hCapName, g_hCapDesc;
static HWND g_hCapCompress, g_hCapStart, g_hCapRefresh;

// Apply controls
static HWND g_hAppWimPath, g_hAppIndex, g_hAppDstVol;
static HWND g_hAppStart, g_hAppRefresh;

// Info controls
static HWND g_hInfWimPath, g_hInfStart, g_hInfResult;

// Pages (groups of controls per tab)
static std::vector<HWND> g_capControls;
static std::vector<HWND> g_appControls;
static std::vector<HWND> g_infControls;
// Labels stored separately
static std::vector<HWND> g_capLabels;
static std::vector<HWND> g_appLabels;
static std::vector<HWND> g_infLabels;

static std::atomic<bool> g_running{false};
static std::thread g_workerThread;

// ============================================================
// Utility Functions
// ============================================================

static HWND CreateLabel(HWND parent, const wchar_t* text, int x, int y, int w, int h)
{
    return CreateWindowW(L"STATIC", text,
        WS_CHILD | SS_LEFT,
        x, y, w, h, parent, nullptr, g_hInst, nullptr);
}

static HWND CreateEdit(HWND parent, int id, int x, int y, int w, int h, DWORD style = 0)
{
    return CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
        WS_CHILD | WS_TABSTOP | ES_AUTOHSCROLL | style,
        x, y, w, h, parent, (HMENU)(INT_PTR)id, g_hInst, nullptr);
}

static HWND CreateBtn(HWND parent, const wchar_t* text, int id, int x, int y, int w, int h)
{
    return CreateWindowW(L"BUTTON", text,
        WS_CHILD | WS_TABSTOP | BS_PUSHBUTTON,
        x, y, w, h, parent, (HMENU)(INT_PTR)id, g_hInst, nullptr);
}

static HWND CreateCombo(HWND parent, int id, int x, int y, int w, int h)
{
    return CreateWindowW(L"COMBOBOX", L"",
        WS_CHILD | WS_TABSTOP | CBS_DROPDOWNLIST | WS_VSCROLL,
        x, y, w, h, parent, (HMENU)(INT_PTR)id, g_hInst, nullptr);
}

static void AppendLog(const wchar_t* text)
{
    // Can be called from any thread - post message to main thread
    wchar_t* copy = _wcsdup(text);
    PostMessageW(g_hWnd, WM_APPEND_LOG, 0, (LPARAM)copy);
}

static void AppendLogDirect(const wchar_t* text)
{
    int len = GetWindowTextLengthW(g_hLog);
    SendMessageW(g_hLog, EM_SETSEL, len, len);
    SendMessageW(g_hLog, EM_REPLACESEL, FALSE, (LPARAM)text);
    SendMessageW(g_hLog, EM_REPLACESEL, FALSE, (LPARAM)L"\r\n");
    SendMessageW(g_hLog, EM_SCROLLCARET, 0, 0);
}

static void SetStatus(const wchar_t* text)
{
    SetWindowTextW(g_hStatus, text);
}

static void SetProgress(int percent)
{
    PostMessageW(g_hWnd, WM_SET_PROGRESS, (WPARAM)percent, 0);
}

// ============================================================
// Drive Enumeration
// ============================================================

struct VolumeInfo {
    std::wstring path;      // e.g. "C:\\"
    std::wstring label;
    std::wstring fsType;
    ULARGE_INTEGER totalSize;
    ULARGE_INTEGER freeSize;
    std::wstring display;   // for combobox
};

static std::vector<VolumeInfo> EnumVolumes()
{
    std::vector<VolumeInfo> volumes;
    DWORD drives = GetLogicalDrives();

    for (int i = 0; i < 26; i++)
    {
        if (!(drives & (1 << i))) continue;

        VolumeInfo vi = {};
        wchar_t root[] = { (wchar_t)(L'A' + i), L':', L'\\', 0 };
        vi.path = root;

        UINT type = GetDriveTypeW(root);
        if (type == DRIVE_NO_ROOT_DIR) continue;

        wchar_t label[256] = {};
        wchar_t fs[64] = {};
        if (GetVolumeInformationW(root, label, 256, nullptr, nullptr, nullptr, fs, 64))
        {
            vi.label = label;
            vi.fsType = fs;
        }

        GetDiskFreeSpaceExW(root, nullptr, &vi.totalSize, &vi.freeSize);

        double totalGB = (double)vi.totalSize.QuadPart / (1024.0 * 1024.0 * 1024.0);
        double freeGB = (double)vi.freeSize.QuadPart / (1024.0 * 1024.0 * 1024.0);

        wchar_t display[512];
        const wchar_t* typeStr = L"";
        switch (type) {
            case DRIVE_REMOVABLE: typeStr = L"USB"; break;
            case DRIVE_FIXED:     typeStr = L"HDD"; break;
            case DRIVE_REMOTE:    typeStr = L"NET"; break;
            case DRIVE_CDROM:     typeStr = L"CD";  break;
            case DRIVE_RAMDISK:   typeStr = L"RAM"; break;
        }

        swprintf_s(display, L"%c: [%s] %s (%s) %.1fGB / %.1fGB",
            L'A' + i, typeStr, label, fs, freeGB, totalGB);

        vi.display = display;
        volumes.push_back(vi);
    }

    return volumes;
}

static void PopulateVolumeCombo(HWND combo)
{
    SendMessageW(combo, CB_RESETCONTENT, 0, 0);
    auto volumes = EnumVolumes();
    for (auto& v : volumes)
    {
        int idx = (int)SendMessageW(combo, CB_ADDSTRING, 0, (LPARAM)v.display.c_str());
        // Store drive letter path
        wchar_t* data = _wcsdup(v.path.c_str());
        SendMessageW(combo, CB_SETITEMDATA, idx, (LPARAM)data);
    }
    if (volumes.size() > 0)
        SendMessageW(combo, CB_SETCURSEL, 0, 0);
}

static std::wstring GetSelectedVolumePath(HWND combo)
{
    int sel = (int)SendMessageW(combo, CB_GETCURSEL, 0, 0);
    if (sel == CB_ERR) return L"";
    wchar_t* data = (wchar_t*)SendMessageW(combo, CB_GETITEMDATA, sel, 0);
    if (data) return data;
    return L"";
}

// ============================================================
// File Browse Dialogs
// ============================================================

static std::wstring BrowseWimSave(HWND owner)
{
    wchar_t path[MAX_PATH] = L"image.wim";
    OPENFILENAMEW ofn = {};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = owner;
    ofn.lpstrFilter = L"WIM Files (*.wim)\0*.wim\0All Files (*.*)\0*.*\0";
    ofn.lpstrFile = path;
    ofn.nMaxFile = MAX_PATH;
    ofn.Flags = OFN_OVERWRITEPROMPT | OFN_PATHMUSTEXIST;
    ofn.lpstrDefExt = L"wim";
    if (GetSaveFileNameW(&ofn))
        return path;
    return L"";
}

static std::wstring BrowseWimOpen(HWND owner)
{
    wchar_t path[MAX_PATH] = {};
    OPENFILENAMEW ofn = {};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = owner;
    ofn.lpstrFilter = L"WIM Files (*.wim)\0*.wim\0ESD Files (*.esd)\0*.esd\0All Files (*.*)\0*.*\0";
    ofn.lpstrFile = path;
    ofn.nMaxFile = MAX_PATH;
    ofn.Flags = OFN_FILEMUSTEXIST;
    if (GetOpenFileNameW(&ofn))
        return path;
    return L"";
}

// ============================================================
// DISM Execution
// ============================================================

static bool RunDism(const std::wstring& args)
{
    AppendLog((L"[DISM] dism.exe " + args).c_str());

    SECURITY_ATTRIBUTES sa = {};
    sa.nLength = sizeof(sa);
    sa.bInheritHandle = TRUE;

    HANDLE hReadPipe = nullptr, hWritePipe = nullptr;
    if (!CreatePipe(&hReadPipe, &hWritePipe, &sa, 0))
    {
        AppendLog(L"[ERROR] Failed to create pipe");
        return false;
    }
    SetHandleInformation(hReadPipe, HANDLE_FLAG_INHERIT, 0);

    std::wstring cmdline = L"dism.exe " + args;

    STARTUPINFOW si = {};
    si.cb = sizeof(si);
    si.dwFlags = STARTF_USESTDHANDLES | STARTF_USESHOWWINDOW;
    si.hStdOutput = hWritePipe;
    si.hStdError = hWritePipe;
    si.wShowWindow = SW_HIDE;

    PROCESS_INFORMATION pi = {};

    BOOL created = CreateProcessW(
        nullptr, &cmdline[0], nullptr, nullptr, TRUE,
        CREATE_NO_WINDOW, nullptr, nullptr, &si, &pi);

    CloseHandle(hWritePipe);

    if (!created)
    {
        AppendLog(L"[ERROR] Failed to start dism.exe");
        CloseHandle(hReadPipe);
        return false;
    }

    // Read output
    char buffer[4096];
    DWORD bytesRead;
    std::string lineBuffer;

    while (ReadFile(hReadPipe, buffer, sizeof(buffer) - 1, &bytesRead, nullptr) && bytesRead > 0)
    {
        if (!g_running) break;

        buffer[bytesRead] = '\0';
        lineBuffer += buffer;

        // Process complete lines
        size_t pos;
        while ((pos = lineBuffer.find('\n')) != std::string::npos)
        {
            std::string line = lineBuffer.substr(0, pos);
            lineBuffer.erase(0, pos + 1);

            // Trim \r
            if (!line.empty() && line.back() == '\r')
                line.pop_back();

            if (line.empty()) continue;

            // Convert to wide
            int wlen = MultiByteToWideChar(CP_OEMCP, 0, line.c_str(), -1, nullptr, 0);
            std::wstring wline(wlen, 0);
            MultiByteToWideChar(CP_OEMCP, 0, line.c_str(), -1, &wline[0], wlen);
            wline.resize(wcslen(wline.c_str()));

            // Parse progress percentage (e.g. "[===  25.0%  ===]" or just "25.0%")
            // DISM outputs progress like: [==== 10.0% ====]
            for (size_t i = 0; i < wline.size(); i++)
            {
                if (iswdigit(wline[i]) || wline[i] == L'.')
                {
                    // Try to parse a number followed by %
                    wchar_t* end = nullptr;
                    double val = wcstod(&wline[i], &end);
                    if (end && *end == L'%' && val >= 0.0 && val <= 100.0)
                    {
                        SetProgress((int)val);
                        break;
                    }
                }
            }

            AppendLog(wline.c_str());
        }

        // Also handle \r for progress updates (DISM uses \r for progress overwrite)
        while ((pos = lineBuffer.find('\r')) != std::string::npos)
        {
            std::string line = lineBuffer.substr(0, pos);
            lineBuffer.erase(0, pos + 1);

            if (line.empty()) continue;

            int wlen = MultiByteToWideChar(CP_OEMCP, 0, line.c_str(), -1, nullptr, 0);
            std::wstring wline(wlen, 0);
            MultiByteToWideChar(CP_OEMCP, 0, line.c_str(), -1, &wline[0], wlen);
            wline.resize(wcslen(wline.c_str()));

            for (size_t i = 0; i < wline.size(); i++)
            {
                if (iswdigit(wline[i]) || wline[i] == L'.')
                {
                    wchar_t* end = nullptr;
                    double val = wcstod(&wline[i], &end);
                    if (end && *end == L'%' && val >= 0.0 && val <= 100.0)
                    {
                        SetProgress((int)val);
                        break;
                    }
                }
            }
        }
    }

    // Process remaining data
    if (!lineBuffer.empty())
    {
        int wlen = MultiByteToWideChar(CP_OEMCP, 0, lineBuffer.c_str(), -1, nullptr, 0);
        std::wstring wline(wlen, 0);
        MultiByteToWideChar(CP_OEMCP, 0, lineBuffer.c_str(), -1, &wline[0], wlen);
        wline.resize(wcslen(wline.c_str()));
        AppendLog(wline.c_str());
    }

    WaitForSingleObject(pi.hProcess, INFINITE);

    DWORD exitCode = 1;
    GetExitCodeProcess(pi.hProcess, &exitCode);

    CloseHandle(pi.hProcess);
    CloseHandle(pi.hThread);
    CloseHandle(hReadPipe);

    if (exitCode == 0)
    {
        AppendLog(L"[OK] Operation completed successfully.");
        SetProgress(100);
    }
    else
    {
        wchar_t msg[128];
        swprintf_s(msg, L"[ERROR] DISM exited with code: %lu", exitCode);
        AppendLog(msg);
    }

    return exitCode == 0;
}

// ============================================================
// Operations (run in worker thread)
// ============================================================

struct CaptureParams {
    std::wstring srcPath;
    std::wstring dstPath;
    std::wstring name;
    std::wstring desc;
    bool compress;
};

struct ApplyParams {
    std::wstring wimPath;
    int index;
    std::wstring dstPath;
};

struct InfoParams {
    std::wstring wimPath;
};

static void WorkerCapture(CaptureParams p)
{
    g_running = true;
    SetStatus(L"Capturing image...");
    SetProgress(0);

    std::wstring compType = p.compress ? L"maximum" : L"fast";
    std::wstring args = L"/Capture-Image"
        L" /ImageFile:\"" + p.dstPath + L"\""
        L" /CaptureDir:\"" + p.srcPath + L"\""
        L" /Name:\"" + p.name + L"\""
        L" /Description:\"" + p.desc + L"\""
        L" /Compress:" + compType;

    bool ok = RunDism(args);

    PostMessageW(g_hWnd, WM_OPERATION_DONE, ok ? 1 : 0, 0);
    g_running = false;
}

static void WorkerApply(ApplyParams p)
{
    g_running = true;
    SetStatus(L"Applying image...");
    SetProgress(0);

    wchar_t indexStr[16];
    swprintf_s(indexStr, L"%d", p.index);

    std::wstring args = L"/Apply-Image"
        L" /ImageFile:\"" + p.wimPath + L"\""
        L" /Index:" + indexStr +
        L" /ApplyDir:\"" + p.dstPath + L"\"";

    bool ok = RunDism(args);

    PostMessageW(g_hWnd, WM_OPERATION_DONE, ok ? 1 : 0, 0);
    g_running = false;
}

static void WorkerInfo(InfoParams p)
{
    g_running = true;
    SetStatus(L"Getting image info...");
    SetProgress(0);

    std::wstring args = L"/Get-ImageInfo /ImageFile:\"" + p.wimPath + L"\"";
    RunDism(args);

    PostMessageW(g_hWnd, WM_OPERATION_DONE, 1, 0);
    g_running = false;
}

// ============================================================
// UI Helpers
// ============================================================

static std::wstring GetWindowStr(HWND hwnd)
{
    int len = GetWindowTextLengthW(hwnd);
    if (len == 0) return L"";
    std::wstring str(len + 1, 0);
    GetWindowTextW(hwnd, &str[0], len + 1);
    str.resize(len);
    return str;
}

static void EnableOperationUI(bool enable)
{
    EnableWindow(g_hCapStart, enable);
    EnableWindow(g_hAppStart, enable);
    EnableWindow(g_hInfStart, enable);
}

static void ShowTabPage(int page)
{
    auto showGroup = [](std::vector<HWND>& ctrls, std::vector<HWND>& labels, bool show) {
        int sw = show ? SW_SHOW : SW_HIDE;
        for (auto h : ctrls) ShowWindow(h, sw);
        for (auto h : labels) ShowWindow(h, sw);
    };

    showGroup(g_capControls, g_capLabels, page == TAB_CAPTURE);
    showGroup(g_appControls, g_appLabels, page == TAB_APPLY);
    showGroup(g_infControls, g_infLabels, page == TAB_INFO);
}

// ============================================================
// Create UI Controls
// ============================================================

static HFONT CreateAppFont()
{
    return CreateFontW(
        -14, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
        DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
        CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI");
}

static void SetFontRecursive(HWND parent, HFONT font)
{
    SendMessageW(parent, WM_SETFONT, (WPARAM)font, TRUE);
    HWND child = GetWindow(parent, GW_CHILD);
    while (child)
    {
        SendMessageW(child, WM_SETFONT, (WPARAM)font, TRUE);
        child = GetWindow(child, GW_HWNDNEXT);
    }
}

static void CreateControls(HWND hWnd)
{
    HFONT hFont = CreateAppFont();

    // ---- Tab Control ----
    g_hTab = CreateWindowW(WC_TABCONTROLW, L"",
        WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS,
        10, 10, APP_WIDTH - 30, 280,
        hWnd, (HMENU)IDC_TAB, g_hInst, nullptr);

    TCITEMW tie = {};
    tie.mask = TCIF_TEXT;
    tie.pszText = (LPWSTR)L"Capture (Image)";
    TabCtrl_InsertItem(g_hTab, 0, &tie);
    tie.pszText = (LPWSTR)L"Apply (Deploy)";
    TabCtrl_InsertItem(g_hTab, 1, &tie);
    tie.pszText = (LPWSTR)L"Image Info";
    TabCtrl_InsertItem(g_hTab, 2, &tie);

    // Offsets inside tab
    int tx = 25, ty = 45;
    int lw = 120, ew = 420, rh = 28, gap = 34;

    // ============ Capture Tab ============
    {
        int y = ty;
        HWND lbl;

        lbl = CreateLabel(hWnd, L"Source Volume:", tx, y + 3, lw, 20);
        g_capLabels.push_back(lbl);
        g_hCapSrcVol = CreateCombo(hWnd, IDC_CAP_SRCVOL, tx + lw, y, ew, 200);
        g_capControls.push_back(g_hCapSrcVol);
        g_hCapRefresh = CreateBtn(hWnd, L"Refresh", IDC_CAP_REFRESH, tx + lw + ew + 5, y, 70, rh);
        g_capControls.push_back(g_hCapRefresh);

        y += gap;
        lbl = CreateLabel(hWnd, L"Save WIM to:", tx, y + 3, lw, 20);
        g_capLabels.push_back(lbl);
        g_hCapDstPath = CreateEdit(hWnd, IDC_CAP_DSTPATH, tx + lw, y, ew, rh);
        g_capControls.push_back(g_hCapDstPath);
        HWND btn = CreateBtn(hWnd, L"Browse", IDC_CAP_DSTBROWSE, tx + lw + ew + 5, y, 70, rh);
        g_capControls.push_back(btn);

        y += gap;
        lbl = CreateLabel(hWnd, L"Image Name:", tx, y + 3, lw, 20);
        g_capLabels.push_back(lbl);
        g_hCapName = CreateEdit(hWnd, IDC_CAP_NAME, tx + lw, y, ew, rh);
        g_capControls.push_back(g_hCapName);
        SetWindowTextW(g_hCapName, L"Windows Image");

        y += gap;
        lbl = CreateLabel(hWnd, L"Description:", tx, y + 3, lw, 20);
        g_capLabels.push_back(lbl);
        g_hCapDesc = CreateEdit(hWnd, IDC_CAP_DESC, tx + lw, y, ew, rh);
        g_capControls.push_back(g_hCapDesc);
        SetWindowTextW(g_hCapDesc, L"Captured by WinPE Image Tool");

        y += gap;
        g_hCapCompress = CreateWindowW(L"BUTTON", L"Maximum Compression",
            WS_CHILD | WS_TABSTOP | BS_AUTOCHECKBOX,
            tx + lw, y, 200, rh, hWnd, (HMENU)IDC_CAP_COMPRESS, g_hInst, nullptr);
        g_capControls.push_back(g_hCapCompress);
        SendMessageW(g_hCapCompress, BM_SETCHECK, BST_CHECKED, 0);

        y += gap + 5;
        g_hCapStart = CreateBtn(hWnd, L"Start Capture", IDC_CAP_START, tx + lw, y, 150, 32);
        g_capControls.push_back(g_hCapStart);
    }

    // ============ Apply Tab ============
    {
        int y = ty;
        HWND lbl;

        lbl = CreateLabel(hWnd, L"WIM File:", tx, y + 3, lw, 20);
        g_appLabels.push_back(lbl);
        g_hAppWimPath = CreateEdit(hWnd, IDC_APP_WIMPATH, tx + lw, y, ew, rh);
        g_appControls.push_back(g_hAppWimPath);
        HWND btn = CreateBtn(hWnd, L"Browse", IDC_APP_WIMBROWSE, tx + lw + ew + 5, y, 70, rh);
        g_appControls.push_back(btn);

        y += gap;
        lbl = CreateLabel(hWnd, L"Image Index:", tx, y + 3, lw, 20);
        g_appLabels.push_back(lbl);
        g_hAppIndex = CreateEdit(hWnd, IDC_APP_INDEX, tx + lw, y, 60, rh, ES_NUMBER);
        g_appControls.push_back(g_hAppIndex);
        SetWindowTextW(g_hAppIndex, L"1");

        y += gap;
        lbl = CreateLabel(hWnd, L"Apply to Volume:", tx, y + 3, lw, 20);
        g_appLabels.push_back(lbl);
        g_hAppDstVol = CreateCombo(hWnd, IDC_APP_DSTVOL, tx + lw, y, ew, 200);
        g_appControls.push_back(g_hAppDstVol);
        g_hAppRefresh = CreateBtn(hWnd, L"Refresh", IDC_APP_REFRESH, tx + lw + ew + 5, y, 70, rh);
        g_appControls.push_back(g_hAppRefresh);

        y += gap + 5;

        // Warning label
        lbl = CreateLabel(hWnd, L"WARNING: All data on the target volume will be overwritten!",
            tx + lw, y, ew, 20);
        g_appLabels.push_back(lbl);

        y += gap;
        g_hAppStart = CreateBtn(hWnd, L"Start Apply", IDC_APP_START, tx + lw, y, 150, 32);
        g_appControls.push_back(g_hAppStart);
    }

    // ============ Info Tab ============
    {
        int y = ty;
        HWND lbl;

        lbl = CreateLabel(hWnd, L"WIM File:", tx, y + 3, lw, 20);
        g_infLabels.push_back(lbl);
        g_hInfWimPath = CreateEdit(hWnd, IDC_INF_WIMPATH, tx + lw, y, ew, rh);
        g_infControls.push_back(g_hInfWimPath);
        HWND btn = CreateBtn(hWnd, L"Browse", IDC_INF_WIMBROWSE, tx + lw + ew + 5, y, 70, rh);
        g_infControls.push_back(btn);

        y += gap;
        g_hInfStart = CreateBtn(hWnd, L"Get Info", IDC_INF_START, tx + lw, y, 150, 32);
        g_infControls.push_back(g_hInfStart);
    }

    // ============ Common: Log, Progress, Status ============
    int logTop = 300;

    CreateLabel(hWnd, L"Log:", 10, logTop, 50, 20);

    g_hLog = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
        WS_CHILD | WS_VISIBLE | WS_VSCROLL |
        ES_MULTILINE | ES_READONLY | ES_AUTOVSCROLL,
        10, logTop + 20, APP_WIDTH - 30, 170,
        hWnd, (HMENU)IDC_LOG, g_hInst, nullptr);

    g_hProgress = CreateWindowW(PROGRESS_CLASSW, L"",
        WS_CHILD | WS_VISIBLE | PBS_SMOOTH,
        10, logTop + 195, APP_WIDTH - 30, 22,
        hWnd, (HMENU)IDC_PROGRESS, g_hInst, nullptr);
    SendMessageW(g_hProgress, PBM_SETRANGE, 0, MAKELPARAM(0, 100));

    g_hStatus = CreateLabel(hWnd, L"Ready", 10, logTop + 222, APP_WIDTH - 30, 20);
    ShowWindow(g_hStatus, SW_SHOW);

    // ---- Apply font ----
    SetFontRecursive(hWnd, hFont);

    // ---- Populate combos ----
    PopulateVolumeCombo(g_hCapSrcVol);
    PopulateVolumeCombo(g_hAppDstVol);

    // ---- Show initial tab ----
    ShowTabPage(TAB_CAPTURE);
}

// ============================================================
// Command Handlers
// ============================================================

static void OnStartCapture()
{
    if (g_running) return;

    CaptureParams p;
    p.srcPath = GetSelectedVolumePath(g_hCapSrcVol);
    p.dstPath = GetWindowStr(g_hCapDstPath);
    p.name = GetWindowStr(g_hCapName);
    p.desc = GetWindowStr(g_hCapDesc);
    p.compress = (SendMessageW(g_hCapCompress, BM_GETCHECK, 0, 0) == BST_CHECKED);

    if (p.srcPath.empty()) {
        MessageBoxW(g_hWnd, L"Please select a source volume.", APP_TITLE, MB_ICONWARNING);
        return;
    }
    if (p.dstPath.empty()) {
        MessageBoxW(g_hWnd, L"Please specify a destination WIM file path.", APP_TITLE, MB_ICONWARNING);
        return;
    }
    if (p.name.empty()) {
        p.name = L"Windows Image";
    }

    EnableOperationUI(false);

    if (g_workerThread.joinable()) g_workerThread.join();
    g_workerThread = std::thread(WorkerCapture, p);
}

static void OnStartApply()
{
    if (g_running) return;

    ApplyParams p;
    p.wimPath = GetWindowStr(g_hAppWimPath);
    p.dstPath = GetSelectedVolumePath(g_hAppDstVol);

    std::wstring indexStr = GetWindowStr(g_hAppIndex);
    p.index = _wtoi(indexStr.c_str());
    if (p.index < 1) p.index = 1;

    if (p.wimPath.empty()) {
        MessageBoxW(g_hWnd, L"Please select a WIM file.", APP_TITLE, MB_ICONWARNING);
        return;
    }
    if (p.dstPath.empty()) {
        MessageBoxW(g_hWnd, L"Please select a target volume.", APP_TITLE, MB_ICONWARNING);
        return;
    }

    // Confirmation
    wchar_t msg[512];
    swprintf_s(msg, L"Apply image index %d to %s?\n\nAll existing data on the target will be overwritten!",
        p.index, p.dstPath.c_str());
    if (MessageBoxW(g_hWnd, msg, APP_TITLE, MB_YESNO | MB_ICONWARNING | MB_DEFBUTTON2) != IDYES)
        return;

    EnableOperationUI(false);

    if (g_workerThread.joinable()) g_workerThread.join();
    g_workerThread = std::thread(WorkerApply, p);
}

static void OnStartInfo()
{
    if (g_running) return;

    InfoParams p;
    p.wimPath = GetWindowStr(g_hInfWimPath);

    if (p.wimPath.empty()) {
        MessageBoxW(g_hWnd, L"Please select a WIM file.", APP_TITLE, MB_ICONWARNING);
        return;
    }

    EnableOperationUI(false);

    if (g_workerThread.joinable()) g_workerThread.join();
    g_workerThread = std::thread(WorkerInfo, p);
}

// ============================================================
// Window Procedure
// ============================================================

static LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg)
    {
    case WM_CREATE:
        CreateControls(hWnd);
        return 0;

    case WM_NOTIFY:
    {
        auto nmhdr = (NMHDR*)lParam;
        if (nmhdr->idFrom == IDC_TAB && nmhdr->code == TCN_SELCHANGE)
        {
            int page = TabCtrl_GetCurSel(g_hTab);
            ShowTabPage(page);
        }
        break;
    }

    case WM_COMMAND:
    {
        int id = LOWORD(wParam);
        switch (id)
        {
        // Capture
        case IDC_CAP_REFRESH:
            PopulateVolumeCombo(g_hCapSrcVol);
            break;
        case IDC_CAP_DSTBROWSE:
        {
            auto path = BrowseWimSave(hWnd);
            if (!path.empty()) SetWindowTextW(g_hCapDstPath, path.c_str());
            break;
        }
        case IDC_CAP_START:
            OnStartCapture();
            break;

        // Apply
        case IDC_APP_WIMBROWSE:
        {
            auto path = BrowseWimOpen(hWnd);
            if (!path.empty()) SetWindowTextW(g_hAppWimPath, path.c_str());
            break;
        }
        case IDC_APP_REFRESH:
            PopulateVolumeCombo(g_hAppDstVol);
            break;
        case IDC_APP_START:
            OnStartApply();
            break;

        // Info
        case IDC_INF_WIMBROWSE:
        {
            auto path = BrowseWimOpen(hWnd);
            if (!path.empty()) SetWindowTextW(g_hInfWimPath, path.c_str());
            break;
        }
        case IDC_INF_START:
            OnStartInfo();
            break;
        }
        break;
    }

    case WM_APPEND_LOG:
    {
        wchar_t* text = (wchar_t*)lParam;
        if (text)
        {
            AppendLogDirect(text);
            free(text);
        }
        return 0;
    }

    case WM_SET_PROGRESS:
    {
        int pct = (int)wParam;
        SendMessageW(g_hProgress, PBM_SETPOS, pct, 0);

        wchar_t status[64];
        swprintf_s(status, L"Progress: %d%%", pct);
        SetStatus(status);
        return 0;
    }

    case WM_OPERATION_DONE:
    {
        bool success = (wParam == 1);
        EnableOperationUI(true);
        SetStatus(success ? L"Completed successfully." : L"Operation failed.");

        if (success)
            MessageBoxW(hWnd, L"Operation completed successfully.", APP_TITLE, MB_ICONINFORMATION);
        else
            MessageBoxW(hWnd, L"Operation failed. Check the log for details.", APP_TITLE, MB_ICONERROR);
        return 0;
    }

    case WM_CLOSE:
        if (g_running)
        {
            if (MessageBoxW(hWnd, L"An operation is in progress. Force quit?",
                APP_TITLE, MB_YESNO | MB_ICONWARNING) != IDYES)
                return 0;
            g_running = false;
        }
        DestroyWindow(hWnd);
        return 0;

    case WM_DESTROY:
        if (g_workerThread.joinable())
        {
            g_running = false;
            g_workerThread.join();
        }
        PostQuitMessage(0);
        return 0;
    }

    return DefWindowProcW(hWnd, msg, wParam, lParam);
}

// ============================================================
// WinMain
// ============================================================

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE, LPWSTR, int nCmdShow)
{
    g_hInst = hInstance;

    // Init common controls
    INITCOMMONCONTROLSEX icex = {};
    icex.dwSize = sizeof(icex);
    icex.dwICC = ICC_TAB_CLASSES | ICC_PROGRESS_CLASS | ICC_STANDARD_CLASSES;
    InitCommonControlsEx(&icex);

    // Register window class
    WNDCLASSEXW wc = {};
    wc.cbSize = sizeof(wc);
    wc.style = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
    wc.lpszClassName = L"WinPEImageToolClass";
    wc.hIcon = LoadIcon(nullptr, IDI_APPLICATION);
    RegisterClassExW(&wc);

    // Create main window (non-resizable)
    DWORD style = WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX;
    RECT rc = { 0, 0, APP_WIDTH, APP_HEIGHT };
    AdjustWindowRect(&rc, style, FALSE);

    g_hWnd = CreateWindowExW(0, wc.lpszClassName, APP_TITLE, style,
        CW_USEDEFAULT, CW_USEDEFAULT,
        rc.right - rc.left, rc.bottom - rc.top,
        nullptr, nullptr, hInstance, nullptr);

    ShowWindow(g_hWnd, nCmdShow);
    UpdateWindow(g_hWnd);

    // Message loop
    MSG msg;
    while (GetMessageW(&msg, nullptr, 0, 0))
    {
        if (!IsDialogMessageW(g_hWnd, &msg))
        {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }

    return (int)msg.wParam;
}
