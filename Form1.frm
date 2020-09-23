VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   3000
      Top             =   2100
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "xx"
      Height          =   255
      Left            =   780
      TabIndex        =   0
      Top             =   1200
      Width           =   3075
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Declare a new instance of the clsCPUUsage class
Private m_clsCPUUsage As New clsCPUUsage

Private Sub Timer1_Timer()
   Label1.Caption = m_clsCPUUsage.CurrentCPUUsage
End Sub
