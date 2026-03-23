#pragma once

// 简化版的桌宠情绪枚举，可以用来控制对白/动画。
enum class QingPetMood
{
    Idle,
    Curious,
    Sleepy,
};

// 帮助调试或 UI 显示，返回对应情绪的中文描述。
inline const wchar_t* QingPetMoodToString(QingPetMood mood)
{
    switch (mood)
    {
    case QingPetMood::Curious: return L"好奇";
    case QingPetMood::Sleepy: return L"想睡觉";
    default: return L"安静";
    }
}
