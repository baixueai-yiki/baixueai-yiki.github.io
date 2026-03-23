#pragma once
#include <windows.h>
#include <string>

// 显示聊天输入气泡，完成后会调用 QingChatHandleInput。
void QingShowChatInput(HWND hwndParent);

// 聊天输入完成后给出的文字会通过该函数处理。
void QingChatHandleInput(HWND hwnd, const std::wstring& input);
