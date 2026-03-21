#include "qingPath.h"
#include <windows.h>

std::wstring QingGetExeDir()
{
    wchar_t path[MAX_PATH];
    GetModuleFileNameW(NULL, path, MAX_PATH);

    std::wstring fullPath(path);
    size_t pos = fullPath.find_last_of(L"\\/");
    return fullPath.substr(0, pos);
}

std::wstring QingGetContentPath(const std::wstring& relativePath)
{
    return QingGetExeDir() + L"\\" + relativePath;
}