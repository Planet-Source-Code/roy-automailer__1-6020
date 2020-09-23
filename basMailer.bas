Attribute VB_Name = "basMailer"
Option Explicit

Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long

Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
    
Public Const SC_CLOSE = &HF060&
Public Const MF_BYCOMMAND = &H0&
