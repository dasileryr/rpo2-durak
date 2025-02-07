VERSION 5.00
Begin VB.Form Spravka 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "О программе"
   ClientHeight    =   1920
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1620
      IntegralHeight  =   0   'False
      ItemData        =   "Spravka.frx":0000
      Left            =   120
      List            =   "Spravka.frx":0010
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "Spravka"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub
