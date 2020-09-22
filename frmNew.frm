VERSION 5.00
Begin VB.Form frmNew 
   Caption         =   "Create New Journal"
   ClientHeight    =   2310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2310
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1440
      Width           =   4335
   End
   Begin VB.TextBox txtJournalName 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4335
   End
   Begin VB.Label Label2 
      Caption         =   $"frmNew.frx":0000
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "Name of Journal"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'These set our variables to laod our databse tables into
'Remember to select Microsoft DAO Library from Reference
Dim db As Database
Dim rs As Recordset
Dim ws As Workspace



Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdCreate_Click()

Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\xjournal.mdb")
Set rs = db.OpenRecordset("journals", dbOpenTable)

rs.AddNew
    rs("name") = txtJournalName.Text
    'If a password has mot been entered then 'none' is entered
    'in the password field and a 0 in protected means the journal
    'does not require a password
    If txtPassword = vbNullString Then
        rs("password") = "None"
        rs("protected") = "0"
    Else
    'enters password and sets protected to 1 so journal requires
    'password
        rs("password") = txtPassword
        rs("protected") = "1"
    End If
rs.Update

Set rs = db.OpenRecordset("select * from journals where name = ('" & txtJournalName & "')")

    'Appends the caption of the main form with the journal name
    frmJournal.Caption = "XJournal - " & rs("name")
    'Calls the show_entries function on the main form passing the
    'journal_id to the function
    frmJournal.show_entries (rs("journal_id"))
    Unload Me
    
End Sub

Private Sub Form_Load()
'Disables the create button until text is entered in the name field
cmdCreate.Enabled = False

End Sub

Private Sub txtJournalName_Change()
'If the journal name box is not empty then the create button
'becomes enabled
If txtJournalName <> vbNullString Then
    cmdCreate.Enabled = True
Else
    cmdCreate.Enabled = False
End If

End Sub
