VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmJournal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "XJournal"
   ClientHeight    =   6345
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   10140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   10140
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTodayJump 
      Caption         =   "Jump to Today"
      Height          =   255
      Left            =   7200
      TabIndex        =   10
      Top             =   3120
      Width           =   2655
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   8640
      TabIndex        =   9
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   7920
      TabIndex        =   8
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   375
      Left            =   7200
      TabIndex        =   7
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddEntry 
      Caption         =   "Add Entry"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   5760
      Width           =   6855
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   7200
      TabIndex        =   2
      Top             =   720
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   24510466
      CurrentDate     =   37252
   End
   Begin VB.TextBox txtEntry 
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   4200
      Width           =   6855
   End
   Begin VB.TextBox txtJournal 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   480
      Width           =   6855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Date Selection"
      Height          =   3135
      Left            =   7080
      TabIndex        =   3
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label lblDate 
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Selected Date: "
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuSeperator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "Format"
      Begin VB.Menu mnufontsize 
         Caption         =   "Font Size"
         Begin VB.Menu mnufontsmall 
            Caption         =   "Small"
         End
         Begin VB.Menu mnufontmedium 
            Caption         =   "Medium"
         End
         Begin VB.Menu mnufontlarge 
            Caption         =   "Large"
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmJournal"
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

Dim max As Long
Dim i As Long
Dim jid As Long

Dim entry_date As String


Private Sub cmdAddEntry_Click()
'Creates a new record and enters the entry into the database
rs.AddNew
    rs("journal_id") = jid
    rs("entry_date") = entry_date
    rs("entry_time") = Time
    rs("entry") = txtEntry
rs.Update

'Clears the entry text box
txtEntry.Text = ""

'resets the recordset and lists the entries again
Set rs = db.OpenRecordset("Select * from entries where journal_id = (" & jid & ") AND entry_date = ('" & entry_date & "') order by entry_time")
list

End Sub

Private Sub cmdExit_Click()

'Closes the form
Unload Me

End Sub

Private Sub cmdNew_Click()
'Opens the form to create a new journal
frmNew.Show
End Sub

Private Sub cmdOpen_Click()
'Show the form to select which journal we want to open
frmSelect.Show
End Sub

Private Sub cmdTodayJump_Click()
'Once clicked jumps to the current date
MonthView1.Value = Date
lblDate.Caption = MonthView1.Value
entry_date = MonthView1.Value


Set rs = db.OpenRecordset("Select * from entries where journal_id = (" & jid & ") AND entry_date = ('" & entry_date & "') order by entry_time")
list

End Sub

Private Sub Form_Load()

'This line sets ws as a Workspace. I don't know what it
'does but it's needed
Set ws = DBEngine.Workspaces(0)
'This opens the databse into the workspace
Set db = ws.OpenDatabase(App.Path & "\xjournal.mdb")
'This selects the table 'Journals' from the databse
Set rs = db.OpenRecordset("journals", dbOpenTable)



'This sets the date caption to the current date as soon
'as the form loads
MonthView1.Value = Date
lblDate.Caption = MonthView1.Value

jid = 0

cmdAddEntry.Enabled = False

End Sub

Private Sub mnuExit_Click()
'Unloads all data from the form and then exits
Unload Me
End Sub

Private Sub mnufontlarge_Click()
'Changes the font size depending on toolbar option selected
txtJournal.FontSize = "12"
End Sub

Private Sub mnufontmedium_Click()
txtJournal.FontSize = "10"
End Sub

Private Sub mnufontsmall_Click()
txtJournal.FontSize = "8"
End Sub

Private Sub mnuHelpAbout_Click()
'Show the About form
frmAbout.Show
End Sub

Private Sub mnuNew_Click()
frmNew.Show
End Sub

Private Sub mnuOpen_Click()
'Show the form to select which journal we want to open
frmSelect.Show
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
'This sets the caption to the date selected when the
'MonthView1 object changes
lblDate.Caption = MonthView1.Value

entry_date = MonthView1.Value

Set rs = db.OpenRecordset("Select * from entries where journal_id = (" & jid & ") AND entry_date = ('" & entry_date & "') order by entry_time")
list
End Sub

Function show_entries(journal_id)
'Sets the variable 'jid' as the journal id passed from the 'select' form
jid = journal_id

'sets entry_date as the current date
entry_date = Date

Set rs = db.OpenRecordset("Select * from entries where journal_id = (" & jid & ") AND entry_date = ('" & entry_date & "') order by entry_time")
list

End Function

Private Sub list()
'If there are no records found then 'No entries found...' is displayed
If rs.RecordCount = 0 Then
    txtJournal.Text = "No entries found..."

Else
    'Counts the number of records in the field
    max = rs.RecordCount
    txtJournal = ""
    rs.MoveFirst
    
    'Loops through the recordset until there are no more entries
    While Not rs.EOF
    'Display the time the entry was made and then the entry on a new line
    txtJournal = txtJournal & rs("entry_time") & vbCrLf
    txtJournal = txtJournal & rs("entry") & vbCrLf & vbCrLf
    rs.MoveNext
    Wend
    
End If
    
End Sub

Private Sub txtEntry_Change()

'Checks to see if the entry textbox is empty or not
If (txtEntry <> vbNullString) Then
  'Checks to see if a journal has been selected, 0 is default when form loads
  If jid <> 0 Then
    cmdAddEntry.Enabled = True
  End If
Else
    cmdAddEntry.Enabled = False
End If
    
End Sub
