VERSION 5.00
Begin VB.Form Fact 
   Caption         =   "Factorising"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5130
   Icon            =   "Factor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   5130
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Height          =   3495
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   4935
      Begin VB.ListBox List1 
         Height          =   3180
         ItemData        =   "Factor.frx":000C
         Left            =   120
         List            =   "Factor.frx":000E
         TabIndex        =   7
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1575
      Left            =   2640
      TabIndex        =   4
      Top             =   120
      Width           =   2415
      Begin VB.CommandButton Command1 
         Caption         =   "Go"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Max Number"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2415
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Text            =   "1000"
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Min Number"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Text            =   "1"
         Top             =   240
         Width           =   2175
      End
   End
End
Attribute VB_Name = "Fact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Min = Text1.Text
Max = Text2.Text
For Counter = Min To Max
        Text$ = ""
        If Counter Mod 2 = 0 Then
                'IS EVEN
                I = 1
                While (I < Counter / 2)
                        If Counter Mod I = 0 Then
                                Text$ = Text$ + Str$(Counter / I) + ","
                        End If
                        I = I + 1
                Wend
        End If
        If Counter Mod 2 <> 0 Then
                'IS ODD
                I = 1
                While (I < Counter / 2)
                        If Counter Mod I = 0 Then
                                Text$ = Text$ + Str$(Counter / I) + ","
                        End If
                        I = I + 2
                Wend
        End If
        Text$ = Text$ + "1"
        List1.AddItem (Text$)
Next
End Sub
