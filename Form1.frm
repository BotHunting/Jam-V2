VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13065
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   13065
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Calender"
      Height          =   5415
      Left            =   4200
      TabIndex        =   0
      Top             =   840
      Width           =   5055
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   2280
         Top             =   240
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Close"
         Height          =   495
         Left            =   1800
         TabIndex        =   1
         Top             =   4680
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Year"
         Height          =   375
         Left            =   960
         TabIndex        =   10
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Month"
         Height          =   375
         Left            =   960
         TabIndex        =   9
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Day/Time"
         Height          =   375
         Left            =   960
         TabIndex        =   8
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000004&
         Height          =   375
         Left            =   3120
         TabIndex        =   7
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000004&
         Height          =   375
         Left            =   3120
         TabIndex        =   6
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000A&
         Height          =   375
         Left            =   3240
         TabIndex        =   5
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000004&
         Height          =   375
         Left            =   3120
         TabIndex        =   4
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Time"
         Height          =   375
         Left            =   960
         TabIndex        =   3
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000004&
         Height          =   375
         Left            =   3120
         TabIndex        =   2
         Top             =   3120
         Width           =   735
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
Dim Zubed As Variant
Zubed = Now
Label1.Caption = Format(Zubed, "DDDD")
Label2.Caption = Format(Zubed, "DD")
Label3.Caption = Format(Zubed, "MMMM")
Label4.Caption = Format(Zubed, "YYYY")
Label6.Caption = Time
End Sub
