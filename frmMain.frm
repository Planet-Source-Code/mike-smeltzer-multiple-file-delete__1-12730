VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mass File Delete - Mike Smeltzer"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton cmdNextFile 
      Caption         =   "Next File"
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   2760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame frameFileBox 
      Caption         =   "File Box"
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   1200
         Top             =   960
      End
      Begin VB.ListBox lstFiles 
         Height          =   2040
         ItemData        =   "frmMain.frx":0000
         Left            =   120
         List            =   "frmMain.frx":0002
         OLEDropMode     =   1  'Manual
         TabIndex        =   1
         Top             =   360
         Width           =   6375
      End
   End
   Begin VB.Label lblNote 
      Caption         =   "Note: Just Drag and Drop Files onto the 'File Box'"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   3735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mass File Delete by Mike Smeltzer
'If your have any questions or comments
'please email me at ourgroup@sprint.ca

Private Sub cmdNextFile_Click()
'On a Error goto EndOfList:
On Error GoTo EndOfList
'Remove the top item in the listbox
lstFiles.RemoveItem Text
'Select the top item in the listbox
lstFiles.ListIndex = 0
'Call DeleteFile and say the file is the top item in the listbox
DeleteFile lstFiles.Text
'Exit this Sub
Exit Sub
'On a Error it goes to here, When it gets the error it means it's finished
EndOfList:
'Make a beep noise
Beep
'Enable the listbox
lstFiles.Enabled = True
End Sub

Private Sub lstFiles_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'Start the Loop
For i = 1 To 9999
'On a error just continue
On Error Resume Next
'Add the Drag and Dropped file to the list box
lstFiles.AddItem Data.Files(i)
'Continue the Loop
Next i
End Sub

Private Sub cmdDelete_Click()
'On a error goto MustAddFiles:
On Error GoTo MustAddFiles
'Set the listbox to not being enabled
lstFiles.Enabled = False
'Select the top file in the listbox
lstFiles.ListIndex = 0
'Call DeleteFile and say the file is the top item in the listbox
DeleteFile lstFiles.Text
'Exit the Sub
Exit Sub
'On a Error it goes here, It gets a error because there are no files in the listbox
MustAddFiles:
'Create a Message Box and set it's property to Information
MsgBox "You Must Add Some Files First", vbInformation
End Sub

'This is where it goes when you want to delete files
Sub DeleteFile(file)
'On a error goto FileNotFound:
On Error GoTo FileNotFound
'Delete the File, file = the top item in list box
Kill file
'Simulate a click on the invisible command button named cmdNextFile
cmdNextFile.Value = True
'Exit This Sub
Exit Sub
'Go here when theres a error, It goes here because the file could not de deleted
FileNotFound:
'Create a Message Box with it's property Information
MsgBox "The file " & lstFiles.Text & " could not be deleted. The program will continue deletion with other files.", vbInformation
'Simulate a click on the invisible command button named cmdNextFile
cmdNextFile.Value = True
End Sub

