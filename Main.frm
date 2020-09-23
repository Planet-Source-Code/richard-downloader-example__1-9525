VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Downloader Example"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboURL 
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   2775
   End
   Begin VB.CommandButton cmdDownload 
      Caption         =   "&Download"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin Project1.rkDownload rkDownload1 
      Height          =   1335
      Left            =   360
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1320
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   2355
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOTE: This downloader cannot download binary files."
      Height          =   195
      Index           =   4
      Left            =   360
      TabIndex        =   7
      Top             =   3840
      Width           =   3810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "still be downloaded. Try with another file!"
      Height          =   195
      Index           =   3
      Left            =   360
      TabIndex        =   6
      Top             =   3360
      Width           =   2865
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Second link is around 32K or higher, and can"
      Height          =   195
      Index           =   2
      Left            =   360
      TabIndex        =   5
      Top             =   3120
      Width           =   3195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The first link in the combobox is below 32K"
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   4
      Top             =   2880
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Choose a file from the combo box or enter your own, with either http:// or file:// protocol before file, and press 'Download'."
      Height          =   615
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDownload_Click()
If cboURL.Text = "" Then Exit Sub
Dim FoundFile As Boolean

    ' Initiate download
    rkDownload1.GetData cboURL
    
    ' NOTE:
    ' This example saves your file into: APP.PATH + "\TEMP.HTM"
    ' You can change this path/file in the class.
    
    Do
    DoEvents
        If Dir(App.Path + "\temp.htm") = "" Then
            ' file is not found, still looping?
            FoundFile = False
            
        ElseIf Dir(App.Path + "\temp.htm") <> "" Then
            ' file is found, downloaded OK?
            MsgBox "Transfer OK!" + vbCrLf + "File downloaded... If the file is above 32K" + vbCrLf + "we cannot load it into textbox (overflow).", vbInformation
            FoundFile = True
            Exit Do
        End If
    Loop
    
    Select Case FoundFile
        Case True
            cmdDownload.Enabled = True
            rkDownload1.Caption = "Ready..."
        Case False
            rkDownload1.Caption = "File was not transferd correctly." + vbCrLf + "The file:" & cboURL & " was never retreived."
    End Select

End Sub

Private Sub Form_Load()

    With cboURL
        .AddItem "http://ktrk.tripod.com/index.htm"
        .AddItem "http://books.dreambook.com/swedishmetalnetwork/index.htm"
    End With
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' If we are still in the loop, we must use 'End'.
    End
    
End Sub
