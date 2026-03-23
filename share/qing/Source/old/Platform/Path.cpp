#include "qingPath.h"
#include <windows.h>

// 返回当前可执行文件所在目录，用于拼接资源路径。
std::wstring QingGetExeDir()
{
    wchar_t path[MAX_PATH];
    GetModuleFileNameW(nullptr, path, MAX_PATH);

    std::wstring fullPath(path);
    size_t pos = fullPath.find_last_of(L"\\/");
    return fullPath.substr(0, pos);
}

// 以可执行目录为基础拼接给定的相对路径。
std::wstring QingGetContentPath(const std::wstring& relativePath)
{
    return QingGetExeDir() + L"\\" + relativePath;
}
