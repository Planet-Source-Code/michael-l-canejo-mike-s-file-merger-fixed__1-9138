VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMerger 
   Caption         =   "Mike's File Merger"
   ClientHeight    =   3615
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   4935
   Icon            =   "frmMerge.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   4935
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog MergeOutput 
      Left            =   4320
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Output Path"
      Filter          =   "Output Path"
   End
   Begin MSComDlg.CommonDialog AddFile 
      Left            =   3360
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Add File"
      Filter          =   "All File(*.*)|*.*"
   End
   Begin MSComDlg.CommonDialog AddFolder 
      Left            =   3840
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Add Folder"
      Filter          =   "All Folders"
      InitDir         =   "gg"
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Add Folder"
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Add File"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&MERGE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   3120
      Width           =   1095
   End
   Begin VB.ListBox lstFiles 
      Height          =   2595
      ItemData        =   "frmMerge.frx":0E42
      Left            =   120
      List            =   "frmMerge.frx":0E44
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   360
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   "Files: 0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   0
      Width           =   1335
   End
   Begin VB.Menu menuFile 
      Caption         =   "&File"
      Begin VB.Menu menuAddFile 
         Caption         =   "&Add File"
      End
      Begin VB.Menu menu1AddFolder 
         Caption         =   "Add &Folder"
      End
      Begin VB.Menu menuMergeAll 
         Caption         =   "&Merge Files"
      End
      Begin VB.Menu menuGenerate 
         Caption         =   "&Generate Batch File"
      End
      Begin VB.Menu line6 
         Caption         =   "-"
      End
      Begin VB.Menu menuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu menuHelp 
      Caption         =   "&Help"
      Begin VB.Menu menuAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu menu1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu Line1 
         Caption         =   "-"
      End
      Begin VB.Menu menu1AddFile 
         Caption         =   "A&dd File"
         Shortcut        =   ^A
      End
      Begin VB.Menu menuAddFolder 
         Caption         =   "Add &Folder"
         Shortcut        =   ^F
      End
      Begin VB.Menu menuEditPath 
         Caption         =   "&Edit File Path"
         Shortcut        =   ^E
      End
      Begin VB.Menu Line2 
         Caption         =   "-"
      End
      Begin VB.Menu menuMoveUp 
         Caption         =   "Move &Up"
         Shortcut        =   ^U
      End
      Begin VB.Menu menuMovedown 
         Caption         =   "Move &Down"
         Shortcut        =   ^D
      End
      Begin VB.Menu Line4 
         Caption         =   "-"
      End
      Begin VB.Menu menuRemFile 
         Caption         =   "&Remove File"
         Shortcut        =   ^R
      End
      Begin VB.Menu menuRemAll 
         Caption         =   "Remove All"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu line3 
         Caption         =   "-"
      End
   End
End
Attribute VB_Name = "frmMerger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function CenterForm(TheForm As Form): TheForm.Move (Screen.Width) / 2 - (TheForm.Width) / 2, (Screen.Height) / 2 - (TheForm.Height) / 2: End Function
'The CenterForm Function to place a form in the middle of the user's screen

Private Sub Command1_Click()
On Error GoTo TheElse
If lstFiles.ListCount = 0 Then Exit Sub 'If no files are in the listbox, end sub
Dim x As Integer
StartAgain:
MergeOutput.FileName = "" 'resets the filename
    MergeOutput.ShowSave 'shows the Save common dialog
        If MergeOutput.FileName = "" Then Exit Sub 'If no file is selected or the user clicks cancel, exit sub
TheElse:
On Error GoTo endit 'End of the sub
 Open App.Path & "\MergeFiles.bat" For Output As #1 ' Open the path to write to it

        Print #1, "@echo off"
        Print #1, "echo Generated by: Mike's File Merger"
            For x = 0 To lstFiles.ListCount - 1
            If x = 0 Then Print #1, "Copy /b " & Chr(34) & lstFiles.List(0) & Chr(34) & " " & Chr(34) & MergeOutput.FileName & Chr(34): x = 1
            If lstFiles.ListCount = 1 Then GoTo endit
                Print #1, "Copy /b " & Chr(34) & MergeOutput.FileName & Chr(34) & "+" & Chr(34) & lstFiles.List(x) & Chr(34) & " " & Chr(34) & MergeOutput.FileName & Chr(34)
            Next x
            Print #1, "Merged Successfully!"
            Print #1, "Cls" 'Clears the DOS window so it can be closed
            Print #1, "Exit" 'Close the DOS window after it finishes
            'This generates a batch file to merge all the files fromthe lsitbox into one file
            'These are DOS commands that will merge the files from the listbox in the Batch file
    Close #1
    Shell App.Path & "\MergeFiles.bat", vbMinimizedNoFocus
    'Open the batch file and it will then minimize and merge.
    'When it finishes, it will close automatically
endit:
End Sub

Private Sub Command2_Click()
    End 'Ends the program
End Sub

Private Sub Command3_Click()
    AddFile.FileName = "" 'To clear the filepath
        AddFile.ShowOpen 'Shows the Open common dialog
        If AddFile.FileName = "" Then Exit Sub 'If no file is selected or user clicks cancel, end it
    lstFiles.AddItem AddFile.FileName 'Adds the selected file's path from the common dialog
Label1.Caption = "Files: " & lstFiles.ListCount
End Sub

Private Sub Command4_Click()
On Error GoTo endit
Dim i As Long
    AddFolder.FileName = "" 'To clear the filepath
        AddFolder.ShowSave 'Shows the Save common dialog
        If AddFolder.FileName = "" Then Exit Sub 'If no folder is selected or user clicks cancel, end it
        'adds the first file to listbox
        lstFiles.AddItem Replace(AddFolder.FileName, AddFolder.FileTitle, "") & Dir(p4th & exte)
        'loop so more files are added to list box
        For i = 0 To 99999
            'this detects if it listed all the files
            '     and exits the function so it doesnt keep
            '     going on for a long time
            If i > lstFiles.ListCount Then
            MsgBox "D"
                lstFiles.RemoveItem lstFiles.ListCount
                Exit For
            Else
                'this is where its adding the other files from the loop
                lstFiles.AddItem Replace(AddFolder.FileName, AddFolder.FileTitle, "") & Dir
                End If
            Next i
            
endit: Label1.Caption = "Files: " & lstFiles.ListCount
lstFiles.RemoveItem lstFiles.ListCount - 1
End Sub

Private Sub Form_Load()
    CenterForm Me 'Centers form in middle of screen
End Sub

Private Sub Form_Resize()
Me.Width = 5055: Me.Height = 4365
End Sub

Private Sub lstFiles_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbRightButton Then Me.PopupMenu menu1, 0 'On right click display the Menu Context
End Sub

Private Sub lstFiles_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error GoTo EndDrop 'If it reaches the end of the file list, end it
Dim It As Integer
    For It = 1 To 10000 'Let up this amount of files to be added at once
        lstFiles.AddItem Data.Files(It) 'Add path of dropped file
    Next It 'Ends the for and x loop
EndDrop:
End Sub

Private Sub menu1AddFile_Click()
    Command3_Click
End Sub

Private Sub menu1AddFolder_Click()
    Command4_Click
End Sub

Private Sub menuAbout_Click()
msg = "Mike's File Merger was created by: Mike Canejo" & vbCrLf & "E-mail: Mike_3D@hotmail.com" & vbCrLf & "AOL/AIM: Mike3DD" & vbCrLf & vbCrLf & "Help/Info:" & vbCrLf & "    This program generates a batch file (.bat) that will merge all your selected files from the listbox or if you want to just generate the batch file, click the 'File' menu then the 'Generate Batch File' so you can send your files split and let the users merge them together again by using clicking the batch file!. You can add all the files from a folder, just a file, or you can drag and drop a file into the list box! This is a pretty "
msg = msg & "cool idea I came up with but this idea was originally created by a guy that made a program called 'File Splitter' and he used this method to put the files back together after being split. This method is very effective! You can merge audio files together (.mp3, .wav) and others like Avi's or Mpeg's. Just think of all those Mp3's that you like and can be just one big Mp3 file or say you had a DivX movie that was split into two or more files, you can now make it one complete file." & vbCrLf & vbCrLf & "Please E-mail me your comments or"
msg = msg & "suggestions or anything!" & vbCrLf & "-Mike Canejo"
MsgBox msg
End Sub

Private Sub menuAddFile_Click()
    Command3_Click
End Sub

Private Sub menuAddFolder_Click()
    Command4_Click
End Sub

Private Sub menuEditPath_Click()
Dim NewPath As String
    NewPath$ = InputBox("Edit the selected path:", "Edit Path", lstFiles.List(lstFiles.ListIndex)) 'Will Display the VB Input box
        If NewPath$ = "" Then Exit Sub 'If the user leaves the InputBox field blank, end it
    lstFiles.List(lstFiles.ListIndex) = NewPath$ 'Adds the new path altered by the user
End Sub

Private Sub menuExit_Click()
    End
End Sub

Private Sub menuGenerate_Click()
On Error GoTo TheElse
If lstFiles.ListCount = 0 Then Exit Sub 'If no files are in the listbox, end sub
Dim x As Integer
StartAgain:
MergeOutput.FileName = "" 'resets the filename
    MergeOutput.ShowSave 'shows the Save common dialog
        If MergeOutput.FileName = "" Then Exit Sub 'If no file is selected or the user clicks cancel, exit sub
TheElse:
On Error GoTo endit 'End of the sub
 Open Replace(MergeOutput.FileName, MergeOutput.FileTitle, "") & "MergeFiles.bat" For Output As #1 ' Open the path to write to it

        Print #1, "@echo off"
        Print #1, "echo Generated by: Mike's File Merger"
            For x = 0 To lstFiles.ListCount - 1
            If x = 0 Then Print #1, "Copy /b " & Chr(34) & lstFiles.List(0) & Chr(34) & " " & Chr(34) & MergeOutput.FileName & Chr(34): x = 1
            If lstFiles.ListCount = 1 Then GoTo endit
                Print #1, "Copy /b " & Chr(34) & MergeOutput.FileName & Chr(34) & "+" & Chr(34) & lstFiles.List(x) & Chr(34) & " " & Chr(34) & MergeOutput.FileName & Chr(34)
            Next x
            Print #1, "Merged Successfully!"
            Print #1, "Cls" 'Clears the DOS window so it can be closed
            Print #1, "Exit" 'Close the DOS window after it finishes
            'This generates a batch file to merge all the files fromthe lsitbox into one file
            'These are DOS commands that will merge the files from the listbox in the Batch file
    Close #1
endit:
End Sub

Private Sub menuMergeAll_Click()
    Command1_Click
End Sub

Private Sub menuMovedown_Click()
Dim HoldDown As String
If lstFiles.ListIndex = lstFiles.ListCount - 1 Then Exit Sub '
If lstFiles.ListIndex = 0 And lstFiles.ListCount = 0 Then Exit Sub  'If the item is at the bottom and cant go down further then goto end sub
    HoldDown$ = lstFiles.List(lstFiles.ListIndex) 'Stores the item's string below it later to be placed in the current items string
        lstFiles.List(lstFiles.ListIndex) = lstFiles.List(lstFiles.ListIndex + 1) 'Lets the listitem you selected = the one below it
    lstFiles.List(lstFiles.ListIndex + 1) = HoldDown$ 'Lets the item below the one you selected = the one you selected
    lstFiles.ListIndex = lstFiles.ListIndex + 1 'Select the listitem's new location
End Sub

Private Sub menuMoveUp_Click()
Dim HoldUp As String
If lstFiles.ListIndex = 0 Then Exit Sub 'If the item is on top of the list and cant go up further then goto end sub
    HoldUp$ = lstFiles.List(lstFiles.ListIndex) 'Stores the item's string above it later to be placed in the current items string
        lstFiles.List(lstFiles.ListIndex) = lstFiles.List(lstFiles.ListIndex - 1) 'Lets the listitem you selected = the one above it
    lstFiles.List(lstFiles.ListIndex - 1) = HoldUp$ 'Lets the item above the one you selected = the one you selected
    lstFiles.ListIndex = lstFiles.ListIndex - 1 'Select the listitem's new location
End Sub

Private Sub menuRemAll_Click()
    lstFiles.Clear 'Removes all items in the listbox
Label1.Caption = "Files: 0"
End Sub

Private Sub menuRemFile_Click()
    lstFiles.RemoveItem lstFiles.ListIndex
Label1.Caption = "Files: " & lstFiles.ListCount
End Sub
