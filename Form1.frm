VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Date Faker Setup"
   ClientHeight    =   1575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3375
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1575
   ScaleWidth      =   3375
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnBrowse 
      Caption         =   "Browse"
      Height          =   315
      Left            =   2580
      TabIndex        =   14
      Top             =   840
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   480
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   60
      TabIndex        =   13
      Text            =   "Application to run (path)"
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox txtNewMonth 
      Height          =   285
      Left            =   1380
      TabIndex        =   9
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox txtNewDay 
      Height          =   285
      Left            =   1980
      TabIndex        =   8
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox txtNewYear 
      Height          =   285
      Left            =   2700
      TabIndex        =   7
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtYear 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2700
      TabIndex        =   5
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox txtDay 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1980
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Save"
      Height          =   315
      Left            =   1260
      TabIndex        =   1
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox txtMonth 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1380
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1860
      TabIndex        =   12
      Top             =   480
      Width           =   135
   End
   Begin VB.Label Label5 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2460
      TabIndex        =   11
      Top             =   480
      Width           =   135
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Faked Date:"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   540
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Current Date:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   180
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2460
      TabIndex        =   4
      Top             =   120
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1860
      TabIndex        =   3
      Top             =   120
      Width           =   135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function FileExists(strFile) As String
    'This looks for the file and returns false if it isn't there
    On Error Resume Next
    FileExists = Dir(strFile) <> ""
End Function

Private Sub btnBrowse_Click()
    'Opens the CommonDialog control to get the file to fake
    CommonDialog1.ShowOpen
    txtPath.Text = CommonDialog1.FileName
End Sub

Private Sub btnSave_Click()
On Error Resume Next
If txtNewMonth.Text <> "" And txtNewDay.Text <> "" And txtNewYear.Text <> "" Then
    'Make the temp string into DateTime.Date$ format
    temp = txtNewMonth.Text + "-" + txtNewDay.Text + "-" + txtNewYear.Text
    'Open the file to save preferences to
    Open App.Path + "/datefaker.ini" For Append As 1
    Print #1, txtPath.Text
    Print #1, temp
    Close 1
    'Tell all is A-OK
    result = MsgBox("To change the settings, just delete the datefaker.ini in the same folder as Date Faker and relaunch Date Faker.", vbInformation, "How to change settings")
    Beep
    result = MsgBox("Would you like to launch the application now?", vbYesNo, "Launch Application?")
    'If they say yes...
    If result = 6 Then
        'Open the file and read the important information
        temp = CStr(App.Path + "/datefaker.ini")
        FileLength = FileLen(temp)
        Open (temp) For Input As #1
        'Reads all of the file
        wholefile = Input(FileLength, #1)
        fakedate = Left(Right(wholefile, 12), 10)
        filepath = Left(wholefile, Len(wholefile) - 14)
        currentdate = DateTime.Date$
        DateTime.Date$ = fakedate
        'Load the application
        Shell (filepath)
        DateTime.Date$ = currentdate
        Unload Me
    Else
        Unload Me
    End If
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
temp = CStr(App.Path + "/datefaker.ini")
'If the ini file is not there, then open up the options
If FileExists(temp) = "False" Then
    result = MsgBox("Date Faker - Made by special [k]", vbOKOnly, "Date Faker")
    temp = Split(DateTime.Date$, "-")
    txtMonth.Text = temp(0)
    txtDay.Text = temp(1)
    txtYear.Text = temp(2)
'But if the file IS there, then load it up and launch the app
Else
    'Open the file and read the important information
    FileLength = FileLen(temp)
    Open (temp) For Input As #1
    'Reads all of the file
    wholefile = Input(FileLength, #1)
    fakedate = Left(Right(wholefile, 12), 10)
    filepath = Left(wholefile, Len(wholefile) - 14)
    currentdate = DateTime.Date$
    DateTime.Date$ = fakedate
    Shell (filepath)
    DateTime.Date$ = currentdate
    Unload Me
End If
End Sub
