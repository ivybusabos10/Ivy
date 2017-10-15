VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Student Profile"
   ClientHeight    =   10245
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   20250
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10245
   ScaleWidth      =   20250
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo13 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   1
      Left            =   14520
      Style           =   1  'Simple Combo
      TabIndex        =   37
      Top             =   4080
      Width           =   2655
   End
   Begin VB.ComboBox Combo12 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   0
      Left            =   11640
      Style           =   1  'Simple Combo
      TabIndex        =   36
      Top             =   4080
      Width           =   2655
   End
   Begin VB.ComboBox Combo11 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   1
      Left            =   8760
      Style           =   1  'Simple Combo
      TabIndex        =   35
      Top             =   4080
      Width           =   2655
   End
   Begin VB.ComboBox Combo10 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   0
      Left            =   5880
      Style           =   1  'Simple Combo
      TabIndex        =   34
      Top             =   4080
      Width           =   2655
   End
   Begin VB.ComboBox Combo9 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   1
      Left            =   14520
      Style           =   1  'Simple Combo
      TabIndex        =   33
      Top             =   3120
      Width           =   2655
   End
   Begin VB.ComboBox Combo8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   0
      Left            =   11640
      Style           =   1  'Simple Combo
      TabIndex        =   32
      Top             =   3120
      Width           =   2655
   End
   Begin VB.ComboBox Combo7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   1
      Left            =   8760
      Style           =   1  'Simple Combo
      TabIndex        =   31
      Top             =   3120
      Width           =   2655
   End
   Begin VB.ComboBox Combo6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   0
      Left            =   5880
      Style           =   1  'Simple Combo
      TabIndex        =   30
      Top             =   3120
      Width           =   2655
   End
   Begin VB.ComboBox Combo5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   1
      Left            =   14520
      Style           =   1  'Simple Combo
      TabIndex        =   29
      Top             =   2280
      Width           =   2655
   End
   Begin VB.ComboBox Combo4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   0
      Left            =   11640
      Style           =   1  'Simple Combo
      TabIndex        =   28
      Top             =   2280
      Width           =   2655
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   1
      Left            =   8760
      Style           =   1  'Simple Combo
      TabIndex        =   27
      Top             =   2280
      Width           =   2655
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   0
      Left            =   5880
      Style           =   1  'Simple Combo
      TabIndex        =   26
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   480
      TabIndex        =   19
      Top             =   1800
      Width           =   4335
      Begin VB.CommandButton Command6 
         Caption         =   "Print"
         Height          =   495
         Index           =   4
         Left            =   2880
         TabIndex        =   25
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Cancel"
         Height          =   495
         Index           =   3
         Left            =   1440
         TabIndex        =   24
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Save"
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   23
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Delete"
         Height          =   495
         Index           =   1
         Left            =   3000
         TabIndex        =   22
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Edit"
         Height          =   495
         Left            =   1680
         TabIndex        =   21
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1ht 
      Caption         =   "Profile Picture"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   17640
      TabIndex        =   16
      Top             =   1920
      Width           =   2535
      Begin VB.PictureBox CommonDialog1 
         Height          =   2040
         Left            =   120
         ScaleHeight     =   1980
         ScaleWidth      =   1860
         TabIndex        =   17
         Top             =   480
         Width           =   1920
      End
   End
   Begin VB.CommandButton cmdPic 
      BackColor       =   &H8000000D&
      Caption         =   "Upload Photo"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   17640
      MaskColor       =   &H000000FF&
      TabIndex        =   15
      Top             =   4800
      Width           =   2535
   End
   Begin VB.CommandButton cmdlogout 
      BackColor       =   &H8000000D&
      Caption         =   "Logout"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   17760
      MaskColor       =   &H000000FF&
      TabIndex        =   14
      Top             =   9120
      Width           =   2295
   End
   Begin VB.Frame cmdstudentlist 
      Caption         =   "Student List"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   1320
      TabIndex        =   2
      Top             =   6120
      Width           =   15615
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   615
         Left            =   1680
         Top             =   2280
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\dbase2\Database2.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\dbase2\Database2.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1455
         Left            =   120
         TabIndex        =   38
         Top             =   360
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   2566
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   13
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID NO."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Last Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "First Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Midle Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Course"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Year"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Sec"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Birthday"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Parents or Guardian"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Sex"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Address"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Phone No."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Label Text 
      BackStyle       =   0  'Transparent
      Caption         =   "Midle Name"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   3
      Left            =   12360
      TabIndex        =   39
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pangasinan State University"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Index           =   0
      Left            =   6360
      TabIndex        =   18
      Top             =   0
      Width           =   8655
   End
   Begin VB.Label Text 
      BackStyle       =   0  'Transparent
      Caption         =   "Midle Name"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   13
      Left            =   14760
      TabIndex        =   13
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Text 
      BackStyle       =   0  'Transparent
      Caption         =   "Course"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   12
      Left            =   6720
      TabIndex        =   12
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Text 
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   11
      Left            =   9600
      TabIndex        =   11
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Text 
      BackStyle       =   0  'Transparent
      Caption         =   "Birthday"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   10
      Left            =   14880
      TabIndex        =   10
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Text 
      BackStyle       =   0  'Transparent
      Caption         =   "Parents or Guardian"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   9
      Left            =   6000
      TabIndex        =   9
      Top             =   4560
      Width           =   2655
   End
   Begin VB.Label Text 
      BackStyle       =   0  'Transparent
      Caption         =   "Sex"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   8
      Left            =   9720
      TabIndex        =   8
      Top             =   4560
      Width           =   495
   End
   Begin VB.Label Text 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   7
      Left            =   12240
      TabIndex        =   7
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Text 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Number"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   6
      Left            =   14640
      TabIndex        =   6
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label Text 
      BackStyle       =   0  'Transparent
      Caption         =   "Section"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   2
      Left            =   12120
      TabIndex        =   5
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Text 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   1
      Left            =   9240
      TabIndex        =   4
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Text 
      BackStyle       =   0  'Transparent
      Caption         =   "ID Number"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   0
      Left            =   6480
      TabIndex        =   3
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Lingayen Campus"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   9120
      TabIndex        =   1
      Top             =   1320
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Alvear Street Lingayen, Pangasinan"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   7560
      TabIndex        =   0
      Top             =   840
      Width           =   6255
   End
   Begin VB.Image cmdImage1 
      Height          =   18000
      Left            =   -840
      Picture         =   "form2.frx":0000
      Top             =   -120
      Width           =   24000
   End
End
Attribute VB_Name = "form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private CN As ADODB.Connection
Private RS As ADODB.Recordset
Private CMD As ADODB.Command

Dim SQL As String, ID As Long, NewID As Long, _
list As ListItem, strID As String, stridno As String, strlname As String, strfname As String, _
strmname As String, strcourse As String, stryear As String, _
strsection As String, strbday As String, strparentsorguardian As String, _
strsex As String, straddress As String, strcontact As String
Sub ConnDB()
Set CN = New ADODB.Connection
CN.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Database2.mdb;Persist Security Info=False"
CN.Open
Set CMD = New ADODB.Command
Set CMD.ActiveConnection = CN
CMD.CommandType = adCmdText

End Sub

Sub Loadlistview()
SQL = "Select * from Table1"
CMD.CommandText = SQL
Set RS = CMD.Execute
ListView1.ListItems.Clear
With RS
    Do Until .EOF
Set list = ListView1.ListItems.Add(, , !idno & "")
list.SubItems(1) = !lname & ""
list.SubItems(2) = !fname & ""
list.SubItems(3) = !mname & ""
list.SubItems(4) = !course & ""
list.SubItems(5) = !Year & ""
list.SubItems(6) = !sec & ""
list.SubItems(7) = !bday & ""
list.SubItems(8) = !parentsorguardian & ""
list.SubItems(11) = !sex & ""
list.SubItems(12) = !address & ""
list.SubItems(11) = !contact & ""
list.SubItems(12) = CStr(!ID)
.MoveNext
    Loop
End With
With ListView1
If .ListItems.Count > 0 Then
Set .SelectedItem = .ListItems(1)
ListView1_ItemClick .SelectedItem
End If
End With
Set list = Nothing
Set RS = Nothing
End Sub
Private Sub Form_Load()
ConnDB
Loadlistview
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
With list
    Combo1.Text = .Text
    Combo2.Text = .SubItems(1)
    Combo3.Text = .SubItems(2)
    Combo4.Text = .SubItems(3)
    Combo5.Text = .SubItems(4)
    Combo6.Text = .SubItems(5)
    Combo7.Text = .SubItems(6)
    Combo8.Text = .SubItems(7)
    Combo9.Text = .SubItems(8)
    Combo10.Text = .SubItems(9)
    Combo11.Text = .SubItems(10)
    Combo12.Text = .SubItems(11)
    Combo13.Text = .SubItems(12)
End Sub
