Attribute VB_Name = "Commom_Struction"
Option Explicit
'''''''?没有用到的结构,VB编译会不会识别而不造成赘余

Public Type RECT '模块函数传递Type必须Public定义，下次把所有类型，常量放Public在此模块，再慢慢删去其它模块中的重复Type定义
        Left As Long
         Top As Long
        Right As Long
        Bottom As Long
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type
