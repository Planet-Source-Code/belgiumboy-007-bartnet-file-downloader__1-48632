VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BartNet Downloader"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8640
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   8640
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1680
      Top             =   1920
   End
   Begin VB.CommandButton cmdClean 
      Caption         =   "Clean"
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   4560
      Width           =   1215
   End
   Begin InetCtlsObjects.Inet iDownload 
      Left            =   3953
      Top             =   2033
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin ComctlLib.ListView lDownload 
      Height          =   3975
      Left            =   113
      TabIndex        =   3
      Top             =   473
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   7011
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   "FileName"
         Object.Tag             =   ""
         Text            =   "File Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   "FileURL"
         Object.Tag             =   ""
         Text            =   "File URL"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   "Status"
         Object.Tag             =   ""
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdDownload 
      Caption         =   "Continue Downloading"
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   4553
      Width           =   2535
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   7313
      TabIndex        =   7
      Top             =   4553
      Width           =   1215
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   113
      TabIndex        =   6
      Top             =   4553
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   285
      Left            =   7793
      TabIndex        =   5
      Top             =   113
      Width           =   735
   End
   Begin VB.TextBox txtAdd 
      Height          =   285
      Left            =   113
      TabIndex        =   4
      Text            =   "http://"
      Top             =   113
      Width           =   7575
   End
   Begin ComctlLib.ProgressBar p1 
      Height          =   255
      Left            =   113
      TabIndex        =   0
      Top             =   5273
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   0
      Min             =   1e-4
   End
   Begin VB.Label lblProgress 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   1793
      TabIndex        =   2
      Top             =   5033
      Width           =   6735
   End
   Begin VB.Label Label1 
      Caption         =   "Download Status :"
      Height          =   255
      Left            =   113
      TabIndex        =   1
      Top             =   5033
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''#####################################################''
''##                                                 ##''
''##  Created By BelgiumBoy_007                      ##''
''##                                                 ##''
''##  Visit BartNet @ www.bartnet.be for more Codes  ##''
''##                                                 ##''
''##  Copyright 2003 BartNet Corp.                   ##''
''##                                                 ##''
''#####################################################''

Private Sub cmdAdd_Click()
On Error GoTo err
    If Len(txtAdd.Text) <> 0 And txtAdd.Text <> "http://" And Right(txtAdd.Text, 4) <> ".com" Then
        Dim Item As ListItem
        Dim Tmp1 As String
        Dim Tmp2 As String
        Dim Tmp3 As String
        Dim i As Integer
        Dim Success As Boolean
        
        i = Len(txtAdd.Text) - 1
        
        Success = False
        
        Do Until i = 0
            If Mid(txtAdd.Text, i, 1) = "/" Then
                Tmp1 = Mid(txtAdd.Text, i + 1, Len(txtAdd.Text) - i)
                Success = True
                Exit Do
            Else
                i = i - 1
            End If
        Loop
        
        If Success = False Then GoTo err
        
        Tmp2 = txtAdd.Text
        Tmp3 = "PENDING"
        
        Set Item = lDownload.ListItems.Add(, Tmp2, Tmp1)
        Item.SubItems(1) = Tmp2
        Item.SubItems(2) = Tmp3
        
        SaveList
    Else
        GoTo err
    End If
    
    Exit Sub
    
err:
    MsgBox "Please enter a valid file location." & vbCrLf & vbCrLf & "NOTE : make sure you only enter the file once, it cannot be added twice or more", vbOKOnly + vbCritical, "Error"
End Sub

Private Sub SaveList()
    Dim fso As New FileSystemObject
    Dim strm As TextStream
    Dim Item As ListItem
    
On Error GoTo err
    Set strm = fso.OpenTextFile(App.Path & "\List", ForWriting)
    
    For Each Item In lDownload.ListItems
        strm.WriteLine Item.Text
        strm.WriteLine Item.SubItems(1)
        strm.WriteLine Item.SubItems(2)
    Next
    
    strm.Close
    
    SetProgress
    
    Exit Sub
    
err:
    Set strm = fso.CreateTextFile(App.Path & "\List")
    strm.Close
    
    Dim File As File
    
    Set File = fso.GetFile(App.Path & "\List")
    
    File.Attributes = Hidden + System
    
    SaveList
End Sub

Private Sub cmdClean_Click()
    Dim i As Integer
    
On Error Resume Next
    i = 1
    Do Until i = lDownload.ListItems.Count + 2
        If lDownload.ListItems.Item(i).SubItems(2) = "COMPLETED" Or lDownload.ListItems.Item(i).SubItems(2) = "FAILED" Then
            lDownload.ListItems.Remove (i)
        End If
        
        i = i + 1
    Loop
    i = 1
    Do Until i = lDownload.ListItems.Count + 2
        If lDownload.ListItems.Item(i).SubItems(2) = "COMPLETED" Or lDownload.ListItems.Item(i).SubItems(2) = "FAILED" Then
            lDownload.ListItems.Remove (i)
        End If
        
        i = i + 1
    Loop
    i = 1
    Do Until i = lDownload.ListItems.Count + 2
        If lDownload.ListItems.Item(i).SubItems(2) = "COMPLETED" Or lDownload.ListItems.Item(i).SubItems(2) = "FAILED" Then
            lDownload.ListItems.Remove (i)
        End If
        
        i = i + 1
    Loop
    i = 1
    Do Until i = lDownload.ListItems.Count + 2
        If lDownload.ListItems.Item(i).SubItems(2) = "COMPLETED" Or lDownload.ListItems.Item(i).SubItems(2) = "FAILED" Then
            lDownload.ListItems.Remove (i)
        End If
        
        i = i + 1
    Loop
    
    SaveList
End Sub

Private Sub cmdDownload_Click()
    ContinueDownloading
End Sub

Private Sub ContinueDownloading()
    Dim Item As ListItem
    
    For Each Item In lDownload.ListItems
        If Item.SubItems(2) = "PENDING" Then
            Item.SubItems(2) = "DOWNLOADING"
            cmdDownload.Enabled = False
            If GetInternetFile(iDownload, Item.Key, App.Path) = True Then
                Item.SubItems(2) = "COMPLETED"
            Else
                Item.SubItems(2) = "FAILED"
            End If
            
            cmdDownload.Enabled = True
            
            Exit For
        End If
    Next
    
    SaveList
    
    For Each Item In lDownload.ListItems
        If Item.SubItems(2) = "PENDING" Then
            ContinueDownloading
            
            Exit For
        End If
    Next
End Sub

Private Sub cmdExit_Click()
    If MsgBox("Are you sure you want to exit ?", vbYesNo + vbInformation, "Exit") = vbYes Then
        Unload Me
        End
    End If
End Sub

Private Sub cmdRemove_Click()
On Error Resume Next
    lDownload.ListItems.Remove (lDownload.SelectedItem.Index)
    
    SaveList
End Sub

Private Sub Form_Load()
    lDownload.ColumnHeaders(2).Width = lDownload.Width - lDownload.ColumnHeaders(1).Width - lDownload.ColumnHeaders(3).Width - 950
    
    LoadList
    
    Dim Item As ListItem
    Dim Total As Integer
    Dim Done As Integer
    
    For Each Item In lDownload.ListItems
        Total = Total + 1
        If Item.SubItems(2) = "COMPLETED" Or Item.SubItems(2) = "FAILED" Then Done = Done + 1
    Next
    
    If Done < Total And Total <> 0 Then
        If MsgBox("You have " & Total - Done & " File waiting to be downloaded.  Continue downloading ?", vbYesNo + vbQuestion, "Confirm") = vbYes Then
            Timer1.Enabled = True
        End If
    End If
End Sub

Private Sub LoadList()
    Dim fso As New FileSystemObject
    Dim strm As TextStream
    Dim Item As ListItem
    Dim Tmp1 As String
    Dim Tmp2 As String
    Dim Tmp3 As String
    
On Error GoTo err
    Set strm = fso.OpenTextFile(App.Path & "\List", ForReading)
On Error Resume Next
    lDownload.ListItems.Clear
    
    Do Until strm.AtEndOfStream
        Tmp1 = strm.ReadLine
        Tmp2 = strm.ReadLine
        Tmp3 = strm.ReadLine
        
        Set Item = lDownload.ListItems.Add(, Tmp2, Tmp1)
        Item.SubItems(1) = Tmp2
        Item.SubItems(2) = Tmp3
    Loop
        
    strm.Close
    
    SetProgress
    
    Exit Sub
    
err:
    Set strm = fso.CreateTextFile(App.Path & "\List")
    strm.Close
    
    Dim File As File
    
    Set File = fso.GetFile(App.Path & "\List")
    
    File.Attributes = Hidden + System
    
    LoadList
End Sub

Private Sub SetProgress()
On Error Resume Next
    Dim Item As ListItem
    Dim Total As Integer
    Dim Done As Integer
    
    For Each Item In lDownload.ListItems
        Total = Total + 1
        If Item.SubItems(2) = "COMPLETED" Then Done = Done + 1
    Next
    
    p1.Max = Total
    p1.Min = 0
    p1.Value = Done
    lblProgress.Caption = Done & " of " & Total & " files downloaded."
End Sub

Private Sub iDownload_StateChanged(ByVal State As Integer)
    If iDownload.StillExecuting = False Then ContinueDownloading
End Sub

Private Sub Timer1_Timer()
    ContinueDownloading
    Timer1.Enabled = False
End Sub
