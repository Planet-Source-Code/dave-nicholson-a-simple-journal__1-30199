VERSION 5.00
Begin VB.Form frmSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select a journal"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2760
      Width           =   975
   End
   Begin VB.ListBox lstJournals 
      Height          =   1620
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2400
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "If journal is protected a password is required to open it. Please enter the password in the box below before opening the journal"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   4335
   End
End
Attribute VB_Name = "frmSelect"
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

'Declaires variables for bits and bobs
Dim max As Long
Dim errm As String
Dim i As Long
Dim warn As Long
Dim open_journal As Long
Dim jid As Long


Dim del As String

Dim journal As String
Dim password As String

Private Sub cmdDelete_Click()
'Shows a message box to cinfirm deletion
del = MsgBox("Are you sure you want to delete" & vbCrLf & lstJournals.Text & "?", vbQuestion & vbYesNo, "Delete " & lstJournals.Text & "?")

'If yes then...
If del = vbYes Then
    'Gets the journal id from the 'journals' table
    Set rs = db.OpenRecordset("select * from journals where name = ('" & lstJournals.Text & "')")
    jid = rs("journal_id")
    'Gets entries from 'entries' table with the the journal id
    Set rs = db.OpenRecordset("select * from entries where journal_id =(" & jid & ")")
    'Loops through found records and deletes
    While Not rs.EOF
        rs.Delete
        rs.MoveNext
    Wend
    'Re-Selects the journal from the 'journal' table and deletes it
    Set rs = db.OpenRecordset("select * from journals where name = ('" & lstJournals.Text & "')")
    rs.Delete
    Set rs = db.OpenRecordset("journals", dbOpenTable)
    list
Else
    Exit Sub
End If

End Sub

Private Sub cmdExit_Click()
'Unloads the form
Unload Me
End Sub

Private Sub cmdNew_Click()
frmNew.Show
Unload Me
End Sub

Private Sub cmdOpen_Click()
'Sets these variables as the selected item in the listbox and
'the text in the password box
journal = lstJournals.Text
password = txtPassword.Text

'Selects the record from the table which has the name of the
'journal selected using an SQL string
Set rs = db.OpenRecordset("Select * from journals where name = ('" & journal & "')")

'Checks if the journal is protected. If so checks that the password
'entered is correct
If rs("protected") = 1 Then
    If rs("password") <> password Then
        warn = MsgBox("The password entered was incorrect. Please try again", vbCritical, "Incorrect Password")
        txtPassword.Text = ""
        Exit Sub
    Else
         frmJournal.Caption = "XJournal - " & journal
         frmJournal.show_entries (rs("journal_id"))
         Unload Me
    End If
Else
    frmJournal.Caption = "XJournal - " & journal
    frmJournal.show_entries (rs("journal_id"))
    Unload Me
End If

End Sub

Private Sub Form_Load()

'This line sets ws as a Workspace. I don't know what it
'does but it's needed
Set ws = DBEngine.Workspaces(0)
'This opens the databse into the workspace
Set db = ws.OpenDatabase(App.Path & "\xjournal.mdb")
'This selects the table 'Journals' from the databse
Set rs = db.OpenRecordset("journals", dbOpenTable)

'Hide the open buttong until someone selects a journa; from the list
cmdOpen.Enabled = False
cmdDelete.Enabled = False

'Calls the function to list the journals
list

End Sub

Private Function list()

'If there are no journals in the database then does this
If rs.RecordCount = 0 Then
    Exit Function
Else

'Move to the last record and then first to ensure we get an
'accurate record count
rs.MoveLast
rs.MoveFirst

'Clears the list incase any names are still displayed
lstJournals.Clear

'Counts how many records are in the table and then loops for that
'amount adding the journal names the the lstJournals using the
'.Additem command then moves to the next record
max = rs.RecordCount
For i = 1 To max
    lstJournals.AddItem rs("name")
    rs.MoveNext
Next i
End If
End Function

Private Sub lstJournals_Click()

'This checks that a journal has been selected from the list.
'If True Then the Open button becomes enabled
If lstJournals.SelCount > -1 Then
    cmdOpen.Enabled = True
    cmdDelete.Enabled = True
Else
    cmdOpen.Enabled = False
    cmdDelete.Enabled = False
End If
End Sub

