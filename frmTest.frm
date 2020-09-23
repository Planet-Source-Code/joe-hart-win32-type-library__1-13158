VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Type Library Test"
   ClientHeight    =   1335
   ClientLeft      =   6150
   ClientTop       =   2865
   ClientWidth     =   3810
   LinkTopic       =   "Form1"
   ScaleHeight     =   1335
   ScaleWidth      =   3810
   Begin VB.CommandButton Command1 
      Caption         =   "Place On Top"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'This project is a simple test of the jhwin32 type library.  Notice there
'are no declare or constant statements!  You DO NOT have to distribute
'the type library with your EXE.  The type library is only used to compile
'your program.
'
'There are far too many functions and constants (literally thousands)  in
'the type library to test them, so we are only testing a few.  If you want
'to know what else is in the type library, use the Object Browser (F2) to see
'them all.  There is also a description of what most of them do as well.
'
'Note:  If there is a function in the library that doesn't work as you
'expect it, you can override it by using a declare statement.
'
'All functions that use a structure that contains a string are not
'in the type library because of limitations to VB.
'

Private Sub Command1_Click()
   Static fOnTop As Boolean 'flag
   If fOnTop Then
      'Remove from top
      Call SetWindowPos(Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
      Me.Caption = "Type Library Test"
      Command1.Caption = "Place On Top"
   Else
      'make on top
      Call SetWindowPos(Me.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE)
      'this is the API equivilant of Me.Caption = "I'm On Top!"
      Call SetWindowText(Me.hWnd, "I'm On Top!")
      Command1.Caption = "Remove From Top"
   End If
   'toggle flag
   fOnTop = Not fOnTop
End Sub
