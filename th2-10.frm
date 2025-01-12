VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form purchasebilled 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Purchase edition"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12990
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   12990
   ShowInTaskbar   =   0   'False
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10410
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   1110
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   11670
      TabIndex        =   11
      Top             =   7650
      Width           =   1110
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10440
      TabIndex        =   10
      Top             =   7635
      Width           =   1080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   9090
      TabIndex        =   6
      Top             =   7635
      Width           =   1245
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "th2-10.frx":0000
      Height          =   360
      Left            =   1440
      TabIndex        =   1
      Top             =   225
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "ac_name"
      BoundColumn     =   "ac_code"
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bamini"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   330
      Left            =   7095
      TabIndex        =   2
      Top             =   225
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   81461251
      CurrentDate     =   37974
   End
   Begin vb6projectProject1.purbilled bill1 
      Height          =   5910
      Left            =   210
      TabIndex        =   5
      Top             =   1425
      Width           =   12630
      _ExtentX        =   22278
      _ExtentY        =   10425
      billcols        =   2
      billrows        =   1
      billcolwidth    =   ""
      billcolcaption  =   ""
      billlistcolumn  =   0
      database        =   ""
      table           =   ""
      billlistfield   =   ""
      billboundfield  =   ""
      billcoltext     =   ""
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   11520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   1080
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1050
      MaxLength       =   10
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "Text3"
      Top             =   7515
      Width           =   1440
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "th2-10.frx":0014
      Height          =   360
      Left            =   1425
      TabIndex        =   3
      Top             =   825
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "ac_name"
      BoundColumn     =   "ac_code"
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bamini"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   330
      Left            =   7095
      TabIndex        =   4
      Top             =   885
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   81461251
      CurrentDate     =   37974
   End
   Begin VB.Label Label5 
      Caption         =   "Date  :"
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
      Left            =   6240
      TabIndex        =   13
      Top             =   915
      Width           =   675
   End
   Begin VB.Label Label1 
      Caption         =   "Change to  :"
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
      Left            =   150
      TabIndex        =   12
      Top             =   870
      Width           =   1140
   End
   Begin VB.Label Label4 
      Caption         =   "Total  :"
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
      Height          =   270
      Left            =   270
      TabIndex        =   9
      Top             =   7560
      Width           =   750
   End
   Begin VB.Label Label3 
      Caption         =   "Item  :"
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
      Left            =   195
      TabIndex        =   8
      Top             =   240
      Width           =   870
   End
   Begin VB.Label Label2 
      Caption         =   "Date  :"
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
      Left            =   6225
      TabIndex        =   7
      Top             =   270
      Width           =   675
   End
End
Attribute VB_Name = "purchasebilled"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FORMS As String

Private Sub bill1_LostFocus()
Dim I As Integer
Dim temprow As Integer
Dim tempcol As Integer

Text3 = 0

temprow = bill1.billrow
tempcol = bill1.billcol

For I = 1 To bill1.billrows - 1
bill1.billrow = I
bill1.billcol = 3
If bill1.Billcoltext <> "" Then
If IsNumeric(bill1.Billcoltext) Then
Text3 = Text3 + CSng(bill1.Billcoltext)
End If
End If
Next
Text3 = Format(Val(Text3), "0.000")
bill1.billleftcol = 0
bill1.billrow = temprow
bill1.billcol = tempcol
End Sub

Private Sub Command1_Click()
If Not IsNumeric(Text3) Or Text3 = "" Then
    MsgBox "Invalid Bill amount", vbInformation
    Exit Sub
End If

Dim pro As String, qnty As Single, rt As Currency, rt2 As Currency
Dim I As Integer, sql As String
Dim lrcode As String
Dim curserial As Long
Dim entry As Integer, voucher As Single, rateflag As String
Dim salesac As Integer, freightac As Integer, cashac As Integer
Dim supplier As Integer
Dim actype As String, agtype As String, purchaseac As Integer

Dim d As database, r As Recordset, r0 As Recordset, r1 As Recordset, r2 As Recordset
Dim r4 As Recordset, r3 As Recordset, r5 As Recordset, r6 As Recordset

Set d = OpenDatabase(DBPATH)
Set r = d.OpenRecordset("purchase")
r.Index = "voucher"
Set r1 = d.OpenRecordset("stock")
r1.Index = "item"
Set r2 = d.OpenRecordset("mproduct")
r2.Index = "product_id"
'Set r3 = d.OpenRecordset("purchinfo")
'r3.Index = "serial"
Set r5 = d.OpenRecordset("ledger")
r5.Index = "voucher"
Set r6 = d.OpenRecordset("accounts")
r6.Index = "ac_code"

Set r0 = d.OpenRecordset("company")
If Not r0.EOF Then
    r0.MoveFirst
    purchaseac = r0!purchase
End If

r6.Seek "=", purchaseac
If Not r6.NoMatch Then
    agtype = r6!AC_TYPE
    actype = "J"
End If


pro = DBCombo1.BoundText
Set r4 = d.OpenRecordset("select max(serial) as lastserial from purchinfo")
If Not r4.EOF Then
    If Not IsNull(r4!lastserial) Then
    r4.MoveFirst
    curserial = r4!lastserial + 1
    Else
    curserial = 1
    End If
    r4.Close
Else
    curserial = 1
End If
entry = 1


'BeginTrans
For I = 1 To bill1.billrows - 1
    bill1.billrow = I
    bill1.billcol = 0
    voucher = Val(bill1.Billcoltext)
    
    bill1.billcol = 1
    supplier = Val(bill1.Billcoltext)
        
    bill1.billcol = 3
    If bill1.Billcoltext <> "" Then
            qnty = CSng(bill1.Billcoltext)
    Else
            qnty = 0
    End If
    
    bill1.billcol = 4
    If bill1.Billcoltext <> "" Then
        rt = CSng(bill1.Billcoltext)
    Else
        rt = 0
    End If
    
            r.Seek "=", voucher
            If voucher <> 0 Then
            If Not r.NoMatch And r!qnty <> qnty Or Not r.NoMatch And r!rate <> rt Or Not r.NoMatch And r!supplier <> supplier Or r.NoMatch _
             And r!item <> DBCombo2.BoundText Or Not r.NoMatch And r!Date <> DTPicker2 Then
            r.Edit
            r!item = DBCombo2.BoundText
            r!Date = DTPicker2
            r!qnty = qnty
            r!rate = rt
            r!supplier = supplier
            r!serial = curserial
            r.Update
         
'            r1.Seek "=", pro
'            rateflag = "notfound"
'            If Not r1.NoMatch Then
'
'                r1.Edit
'                r1!qnty = r1!qnty - qnty
'                r1.Update
'            Else
'                r1.AddNew
'                r1!item = pro
'                r1!qnty = qnty
'                r1.Update
'            End If
            
        'ledger entry
        r5.Seek "=", voucher
        If Not r5.NoMatch Then
        r5.Edit
        r5!Date = DTPicker2
        r5!account = purchaseac  'code for sales a/c
        r5!against = supplier
        r5!amount = qnty * rt
        r5!actype = actype
        r5!agtype = agtype
        r5!drcr = "D"
        'r5!narration = Text4
        r5.Update
        End If
        
        
        r5.MoveNext
        If r5!voucher = voucher Then
        r5.Edit
        r5!Date = DTPicker2
        r5!account = supplier
        r5!against = purchaseac  'code for sales a/c
        r5!amount = qnty * rt
        r5!drcr = "C"
        r5!actype = agtype
        r5!agtype = actype
        'r5!narration = Text4
        r5.Update
        End If
        End If
   End If
   
Next


        
'CommitTrans

    r.Close
    r1.Close
    r2.Close
'    r3.Close
    r5.Close
    d.Close

    MsgBox "Entries saved", vbInformation
    Command2_Click

End Sub

Private Sub Command2_Click()
Dim I As Integer
bill1.billrow = 1
For I = 1 To bill1.billrows - 1
    bill1.billcol = 0
    bill1.Billcoltext = ""
    bill1.billcol = 2
    bill1.Billcoltext = ""
    bill1.billcol = 3
    bill1.Billcoltext = ""
    bill1.billcol = 4
    bill1.Billcoltext = ""
    bill1.billcol = 5
    bill1.Billcoltext = ""
    If bill1.billrow < bill1.billrows - 1 Then
    bill1.billrow = bill1.billrow + 1
    End If
Next

Text3 = ""


bill1.billcol = 1
bill1.billrow = 1
bill1.postext1

DBCombo1.SetFocus

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub DTPicker1_LostFocus()
Dim d As database, r As Recordset, r1 As Recordset, sql As String
sql = "select * from purchase where item='" + DBCombo1.BoundText + "' and "
sql = sql + " purchase.date=#" & DTPicker1.Month & "/" & DTPicker1.day & "/" & DTPicker1.year & "#"
Set d = OpenDatabase(DBPATH)
Set r1 = d.OpenRecordset("accounts")
r1.Index = "ac_code"

DTPicker2 = DTPicker1
DBCombo2 = DBCombo1
Set r = d.OpenRecordset(sql)

bill1.billrows = 1
bill1.billrows = 2
bill1.billrow = 1

If Not r.NoMatch Then
    Do While Not r.EOF
        bill1.billcol = 0
        bill1.Billcoltext = r!voucher
        bill1.billcol = 1
        bill1.Billcoltext = r!supplier
        
        bill1.billcol = 2
        
        r1.Seek "=", r!supplier
        If Not r1.NoMatch Then
            bill1.Billcoltext = r1!ac_name
        End If
        bill1.billcol = 3
        bill1.Billcoltext = Format(r!qnty, "0.000")
        bill1.billcol = 4
        bill1.Billcoltext = Format(r!rate, "0.00")
        bill1.billcol = 5
        bill1.Billcoltext = Format(r!rate * r!qnty, "0.00")
        
        If bill1.billrow < bill1.billrows - 1 Then
        bill1.billrow = bill1.billrow + 1
        Else
        bill1.billrows = bill1.billrows + 1
        bill1.billrow = bill1.billrow + 1
        End If
        r.MoveNext
    Loop
End If
bill1.billcol = 2
bill1.billrow = 1
bill1.postext1
bill1.text1text
End Sub

Private Sub Form_Activate()
Data1.DatabaseName = DBPATH
Data1.RecordSource = "select * from mproduct order by product"
DBCombo1.ListField = "product"
DBCombo1.BoundColumn = "product_id"

Data2.DatabaseName = DBPATH
Data2.RecordSource = "select * from mproduct order by product"
DBCombo2.ListField = "product"
DBCombo2.BoundColumn = "product_id"

If Not Data1.Recordset.EOF Then
    Data1.Recordset.MoveFirst
    DBCombo1 = Data1.Recordset!product_id
    DBCombo2 = Data1.Recordset!product_id
End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 123 Then
    Unload Me
End If

End Sub

Private Sub Form_Load()
Me.Top = MDIForm1.Height / 2 - Me.Height / 2 - 400
Me.Left = MDIForm1.Width / 2 - Me.Width / 2


FORMS = "ONE"

bill1.billrows = 50
bill1.billcols = 6
bill1.billcolwidth = "1000;10;6000;1700;1700;1700;"
bill1.billcolcaption = "||Particulars|Qty|Rate|Amount|"
bill1.database = DBPATH
bill1.table = "select * from accounts where ac_type='J' order by ac_name"
bill1.billlistcolumn = 2
bill1.billlistfield = "ac_name"
bill1.billboundfield = "ac_code"

bill1.billcol = 2
bill1.billleftcol = 0
bill1.billtextpos = 100

Text3 = ""
DTPicker1 = TRANSACTION_DATE


'Text6 = ""
'Text7 = ""
End Sub

Private Sub SSCommand1_Click()

End Sub

Private Sub SSCommand2_Click()
End Sub

Private Sub SSCommand3_Click()

End Sub

