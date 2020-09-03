VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Mencetak dengan Pengaturan Letak"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   1920
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub PrintAlignedText(s As String, Alignment _
As String)
    Select Case Alignment
    Case "Center"
        Printer.CurrentX = (Printer.ScaleWidth - _
        Printer.TextWidth(s)) \ 2
    Case "Left"
        Printer.CurrentX = 0
    Case "Right"
        Printer.CurrentX = Printer.ScaleWidth - _
        Printer.TextWidth(s)
    End Select
    Printer.Print s
    'Gunakan perintah EndDoc jika text ini adalah yang
    'terakhir yang ingin Anda cetak ke dalam kertas...
    Printer.EndDoc
End Sub

Private Sub Command1_Click()
   'Ganti tulisan "Rahmat Putra" dengan teks yang akan
   'Anda cetak.
   'Ganti kata "Center" dengan alignment yang Anda
   'inginkan ("Center", "Left" or "Right")
   Call PrintAlignedText("Rahmat Putra", "Center")
End Sub


