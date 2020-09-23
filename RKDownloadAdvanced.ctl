VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl rkDownload 
   BackColor       =   &H00000000&
   ClientHeight    =   1170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3795
   ScaleHeight     =   1170
   ScaleWidth      =   3795
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   360
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblStat 
      BackColor       =   &H00000000&
      Caption         =   "Ready..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "rkDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Written by Richard of KTRK
' http://ktrk.tripod.com

' Usage:
' strPath must start with either HTTP://  or FILE://, or whichever protocol
' you have planned to use.

' Errors are displayed in the caption property of this class.


'<----------------- For the control itself, RESIZE ------------------------->

'Default Property Values:
Const m_def_ScaleWidth = 0
Const m_def_ScaleHeight = 0
'Property Variables:
Dim m_ScaleWidth As Integer
Dim m_ScaleHeight As Integer
Public Sub GetData(strPath As String)


    ' Remove temporary file if it exist
    If Dir(App.Path + "\temp.htm") <> "" Then Kill App.Path + "\temp.htm"

    ' Initiate download
    UserControl.AsyncRead strPath, vbAsyncTypeFile, "Links", vbAsyncReadForceUpdate

End Sub

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
Dim FileNum As Long

On Error GoTo err:

Select Case AsyncProp.PropertyName
    Case "Links"

        Close FileNum
        FileNum = FreeFile
        Open AsyncProp.Value For Input As FileNum
            Open App.Path + "\temp.htm" For Output As #2
                Print #2, Input(LOF(FileNum), FileNum)      ' save to swap space (read first)
            Close #2
        Close FileNum
End Select

Exit Sub

err:
    lblStat = "Err: (" & err.Number & ") - (" & err.Description & ")"
    
    If Dir(App.Path + "\temp.htm") <> "" Then
        ' delete the file, contains probably nothing, or just
        ' alot of crap.
        Close #2, FileNum
        Kill App.Path + "\temp.htm"
    End If
    
End Sub

Private Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)

If AsyncProp.BytesRead = 0 Then
Else
    ' If we have a value of BytesMax...
    If AsyncProp.BytesMax <> 0 Then
        PB1.Max = AsyncProp.BytesMax
    Else
        PB1.Value = 0
    End If
    
    ' Increment progressbar
    PB1.Value = AsyncProp.BytesRead
    ' Display read progress...
    lblStat = "Reading: " & AsyncProp.BytesRead & " / " & AsyncProp.BytesMax & " ..."
End If





   Select Case AsyncProp.StatusCode
      Case vbAsyncStatusCodeSendingRequest
         lblStat = "Connecting to host..."
      Case vbAsyncStatusCodeEndDownloadData
         lblStat = "Internet transfer complete."
      Case vbAsyncStatusCodeError
         lblStat = "Err: Download aborted, " & AsyncProp.StatusCode
         CancelAsyncRead "Links"
   End Select

End Sub
Public Sub UserControl_Resize()
On Error Resume Next

    With PB1
        If Not .Left = 0 Then .Left = 0
        If Not .Top = 0 Then .Top = 0
        .Width = UserControl.ScaleWidth
    End With
        
        
    With lblStat
        If Not .Left = 0 Then .Left = 0
        If Not .Top = PB1.Top + PB1.Height + 50 Then .Top = PB1.Top + PB1.Height + 50
        .Width = UserControl.ScaleWidth
        .Height = UserControl.ScaleHeight - .Top
   End With
   
   

End Sub
'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_ScaleWidth = PropBag.ReadProperty("ScaleWidth", m_def_ScaleWidth)
    m_ScaleHeight = PropBag.ReadProperty("ScaleHeight", m_def_ScaleHeight)
    lblStat.Caption = PropBag.ReadProperty("Caption", "Ready...")
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ScaleWidth", m_ScaleWidth, m_def_ScaleWidth)
    Call PropBag.WriteProperty("ScaleHeight", m_ScaleHeight, m_def_ScaleHeight)
    Call PropBag.WriteProperty("Caption", lblStat.Caption, "Ready...")
End Sub

Public Property Get ScaleWidth() As Single
    ScaleWidth = UserControl.ScaleWidth
End Property

Public Property Let ScaleWidth(ByVal New_ScaleWidth As Single)
    UserControl.ScaleWidth() = New_ScaleWidth
    PropertyChanged "ScaleWidth"
End Property
Public Property Get ScaleHeight() As Single
    ScaleHeight = UserControl.ScaleHeight
End Property

Public Property Let ScaleHeight(ByVal New_ScaleHeight As Single)
    UserControl.ScaleHeight() = New_ScaleHeight
    PropertyChanged "ScaleHeight"
End Property
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = lblStat.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblStat.Caption() = New_Caption
    PropertyChanged "Caption"
End Property
