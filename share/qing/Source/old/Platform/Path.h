#pragma once
#include <string>

// 获取当前可执行文件所在目录（不包含文件名）。
std::wstring QingGetExeDir();

// 以 exe 目录为基础拼接相对路径，方便加载 Content 下的资源。
std::wstring QingGetContentPath(const std::wstring& relativePath);
