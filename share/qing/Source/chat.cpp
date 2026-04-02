#include "chat.h"
#include <iostream>

// 负责将聊天消息写入日志/输出，当前只是打印到控制台
// 你可以在这里接入网络、AI、或者展示到 UI 上
void SendChat(const char* msg)
{
    std::cout << msg << std::endl;
}
