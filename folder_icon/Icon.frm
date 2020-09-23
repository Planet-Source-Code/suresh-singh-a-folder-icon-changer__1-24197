VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Folder Icon Created by Pradeep Singh"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   8790
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Set Default Windows Icon"
      Height          =   465
      Left            =   5310
      TabIndex        =   11
      Top             =   3960
      Width           =   1545
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   465
      Left            =   7110
      TabIndex        =   10
      Top             =   3960
      Width           =   1545
   End
   Begin VB.Frame Frame2 
      Caption         =   "Selected Folder and Icon"
      Height          =   2310
      Left            =   3645
      TabIndex        =   5
      Top             =   135
      Width           =   4920
      Begin VB.Frame Frame4 
         Caption         =   "Icon"
         Height          =   645
         Left            =   90
         TabIndex        =   8
         Top             =   1350
         Width           =   4695
         Begin VB.Label Label2 
            Height          =   285
            Left            =   90
            TabIndex        =   9
            Top             =   270
            Width           =   4425
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Folder"
         Height          =   645
         Left            =   135
         TabIndex        =   6
         Top             =   405
         Width           =   4650
         Begin VB.Label Label1 
            Height          =   285
            Left            =   90
            TabIndex        =   7
            Top             =   225
            Width           =   4290
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Folder and Icon"
      Height          =   4245
      Left            =   90
      TabIndex        =   1
      Top             =   135
      Width           =   3300
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   270
         TabIndex        =   4
         Top             =   315
         Width           =   2850
      End
      Begin VB.DirListBox Dir1 
         Height          =   1665
         Left            =   270
         TabIndex        =   3
         Top             =   765
         Width           =   2850
      End
      Begin VB.FileListBox File1 
         Height          =   1650
         Left            =   270
         Pattern         =   "*.ico"
         TabIndex        =   2
         Top             =   2475
         Width           =   2850
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change Icon"
      Height          =   465
      Left            =   3600
      TabIndex        =   0
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      Height          =   870
      Left            =   5535
      Top             =   2700
      Width           =   960
   End
   Begin VB.Image Image1 
      Height          =   510
      Left            =   5715
      Top             =   2880
      Width           =   600
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#########################################################################
'#                                                                       #
'#               Created by Pradeep Singh                                #
'#               Created on June 19 2001 at 3:06 AM                      #
'#               mail me@- pradeepsingh10@hotmail.com                    #
'#                                                                       #
'#########################################################################
Public Sub CreateFile(fpath As String, Text As String)
On Error Resume Next
Dim f As Integer
f = FreeFile
Open fpath For Output As #f
Print #f, Text
Close #f
End Sub
Private Sub Command1_Click()
On Error Resume Next
'Create and Save Desktop.ini in that folder
CreateFile Label1.Caption & "\Desktop.ini", "[.ShellClassInfo]" & vbNewLine & _
                                                  "IconFile=" + Label2.Caption & vbNewLine & _
                                                  "IconIndex = 0"

'IMPORTANT        IMPORTANT      IMPORTANT      IMPORTANT

'The main part of this hole proccess to make folder property to system
'without changing the folder attribute to system your custom
'icon cannot be shown.

SetAttr Label1.Caption, vbSystem
MsgBox "Select the folder where you have change the icon and press F5 key", vbInformation, "Important"
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
On Error Resume Next
Kill Label1.Caption + "\Desktop.ini"
SetAttr Label1.Caption, vbHidden
MsgBox "Select the folder and press F5 key", vbInformation, "Important"
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
Label1.Caption = Dir1.Path
End Sub

Private Sub Drive1_Change()
'Update directory
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
On Error Resume Next
Dim picimg As String
Dim IconPath As String
'Simple way to load icon in image control
'Just add file1 path and file1 filename and load into image control
Label2.Caption = File1.Path + "\" + File1.FileName
IconPath = File1.Path + "\" + File1.FileName
Image1.Picture = LoadPicture(IconPath)
Label4.Caption = File1.Path + "\" + File1.FileName
End Sub

