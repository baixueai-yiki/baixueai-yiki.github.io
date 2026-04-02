#include "state.h"
#include "common.h"

// 全局状态缓存，初始为 idle
static PetState g_state = STATE_IDLE;

// 外部接口设置当前桌宠状态
void SetState(PetState state)
{
    g_state = state;
}

// 外部接口获取当前桌宠状态
PetState GetState()
{
    return g_state;
}
