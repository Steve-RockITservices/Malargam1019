VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form Compercent 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Commission Percent"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3945
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   3945
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1485
      TabIndex        =   4
      Top             =   2430
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2310
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   225
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.TextBox TxtPercent 
      Height          =   360
      Left            =   2340
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1455
      Width           =   1365
   End
   Begin MSDBCtls.DBCombo DBCSupplier 
      Bindings        =   "th2-12.frx":0000
      Height          =   345
      Left            =   225
      TabIndex        =   0
      Top             =   825
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   609
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "ac_name"
      BoundColumn     =   "ac_code"
      Text            =   "DBCSupplier"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFC0&
      Height          =   645
      Left            =   120
      Top             =   2325
      Width           =   3705
   End
   Begin VB.Label Label1 
      Caption         =   "Commission (%) :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   1530
      Width           =   1650
   End
   Begin VB.Label Label3 
      Caption         =   "Supplier :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   225
      TabIndex        =   1
      Top             =   420
      Width           =   870
   End
End
Attribute VB_Name = "Compercent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As database
Dim rs As Recordset


Private Sub DBCombo1_Click(Area As Integer)

End Sub

Private Sub CmdSave_Click()
rs.Seek "=", DBCSupplier.BoundText
If Not rs.NoMatch Then
    rs.Edit
    rs!Percent = Val(TxtPercent)
    rs!Date = Date
    rs.Update
    MsgBox "Entry edited", vbInformation
Else
    rs.AddNew
    rs!accode = DBCSupplier.BoundText
    rs!Percent = Val(TxtPercent)
    rs!Date = Date
    rs.Update
    MsgBox "Entry saved", vbInformation
End If
DBCSupplier.SetFocus

End Sub

Private Sub DBCSupplier_LostFocus()
rs.Seek "=", DBCSupplier.BoundText
If Not rs.NoMatch Then
    TxtPercent = rs!Percent
Else
    TxtPercent = ""
End If

End Sub

Private Sub Form_Load()
Data1.DatabaseName = DBPATH
Data1.RecordSource = "select * from accounts where ac_type='J' OR AC_TYPE='M' order by ac_name"
DBCSupplier.ListField = "ac_name" '"product"
DBCSupplier.BoundColumn = "AC_CODE" '"product_id"
TxtPercent = ""
Me.Top = MDIForm1.Height / 2 - Me.Height / 2 - 400
Me.Left = MDIForm1.Width / 2 - Me.Width / 2


Set db = OpenDatabase(DBPATH)
Set rs = db.OpenRecordset("commission")
rs.Index = "accode"

End Sub
