VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "dblist32.ocx"
Begin VB.Form salebilled 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sales edition"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13410
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   13410
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Txttotal1 
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
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   18
      TabStop         =   0   'False
      Text            =   "Text3"
      Top             =   7560
      Width           =   1440
   End
   Begin VB.TextBox Txtfrom 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9960
      TabIndex        =   2
      Text            =   "Txtfrom"
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Txtto 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11880
      TabIndex        =   3
      Text            =   "Txtto"
      Top             =   360
      Width           =   1215
   End
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10905
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   945
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
      Height          =   390
      Left            =   12030
      TabIndex        =   13
      Top             =   7605
      Width           =   1125
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
      Left            =   10695
      TabIndex        =   12
      Top             =   7590
      Width           =   1230
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000040C0&
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
      Height          =   405
      Left            =   9420
      TabIndex        =   6
      Top             =   7590
      Width           =   1170
   End
   Begin vb6projectProject1.salbilled bill1 
      Height          =   6045
      Left            =   165
      TabIndex        =   7
      Top             =   1200
      Width           =   13035
      _ExtentX        =   22992
      _ExtentY        =   10663
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
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "th2-11.frx":0000
      Height          =   345
      Left            =   1695
      TabIndex        =   0
      Top             =   210
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   609
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
      Left            =   7245
      TabIndex        =   1
      Top             =   270
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   95027203
      CurrentDate     =   37974
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   975
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.TextBox Txttotal 
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
      Left            =   3000
      MaxLength       =   10
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "Text3"
      Top             =   7560
      Width           =   1440
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "th2-11.frx":0014
      Height          =   345
      Left            =   1710
      TabIndex        =   4
      Top             =   795
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   609
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
      Left            =   7260
      TabIndex        =   5
      Top             =   795
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   95027203
      CurrentDate     =   37974
   End
   Begin VB.Label Label6 
      Caption         =   "From :"
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
      Left            =   9240
      TabIndex        =   17
      Top             =   360
      Width           =   675
   End
   Begin VB.Label Label7 
      Caption         =   "To :"
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
      Left            =   11280
      TabIndex        =   16
      Top             =   360
      Width           =   675
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
      Left            =   6540
      TabIndex        =   15
      Top             =   810
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
      Left            =   225
      TabIndex        =   14
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
      Left            =   255
      TabIndex        =   11
      Top             =   7575
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
      TabIndex        =   10
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
      Left            =   6525
      TabIndex        =   9
      Top             =   285
      Width           =   675
   End
End
Attribute VB_Name = "salebilled"
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

Txttotal = 0

temprow = bill1.billrow
tempcol = bill1.billcol

For I = 1 To bill1.billrows - 1
bill1.billrow = I
bill1.billcol = 5
If bill1.Billcoltext <> "" Then
If IsNumeric(bill1.Billcoltext) Then
Txttotal = Val(Txttotal) + Val(bill1.Billcoltext)
End If
End If
Next
Txttotal = Format(Val(Txttotal), "0.000")
bill1.billleftcol = 0
bill1.billrow = temprow
bill1.billcol = tempcol
End Sub

Private Sub Command1_Click()
If Not IsNumeric(Txttotal) Or Txttotal = "" Then
    MsgBox "Invalid Bill amount", vbInformation
    Exit Sub
End If

Dim pro As String, qnty As Single, rt As Currency, rt2 As Currency
Dim I As Integer, sql As String
Dim lrcode As String
Dim curserial As Long
Dim entry As Integer, voucher As Single, rateflag As String
Dim salesac As Integer, freightac As Integer, cashac As Integer
Dim customer As Integer
Dim actype As String, agtype As String, purchaseac As Integer

Dim d As database, r As Recordset, r0 As Recordset, r1 As Recordset, r2 As Recordset
Dim r4 As Recordset, r3 As Recordset, r5 As Recordset, r6 As Recordset

Set d = OpenDatabase(DBPATH)
Set r = d.OpenRecordset("sales")
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
    salesac = r0!sales
End If

r6.Seek "=", salesac
If Not r6.NoMatch Then
    agtype = r6!AC_TYPE
    actype = "I"
End If


pro = DBCombo1.BoundText
'Set r4 = d.OpenRecordset("select max(serial) as lastserial from info")
'If Not r4.EOF Then
'    If Not IsNull(r4!lastserial) Then
'    r4.MoveFirst
'    curserial = r4!lastserial + 1
'    Else
'    curserial = 1
'    End If
'    r4.Close
'Else
'    curserial = 1
'End If
entry = 1


'BeginTrans
For I = 1 To bill1.billrows - 1
    bill1.billrow = I
    
            bill1.billcol = 0
            voucher = Val(bill1.Billcoltext)
            
            bill1.billcol = 1
            customer = Val(bill1.Billcoltext)
                
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
            If Not r.NoMatch And r!qnty <> qnty Or Not r.NoMatch And r!rate <> rt Or Not r.NoMatch And r!customer <> customer _
            Or Not r.NoMatch And r!item <> DBCombo2.BoundText Or Not r.NoMatch And r!Date <> DTPicker2 Then
                    r.Edit
                    r!Date = DTPicker2
                    r!item = DBCombo2.BoundText
                    r!qnty = qnty
                    r!rate = rt
                    r!customer = customer
                    r!serial = curserial
                    r.Update
                     
            
                    
                'ledger entry
                    r5.Seek "=", voucher
                    If Not r5.NoMatch Then
                        r5.Edit
                        r5!Date = DTPicker2
                        r5!account = customer 'purchaseac  'code for sales a/c
                        r5!against = salesac 'customer
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
                        r5!account = salesac 'customer
                        r5!against = customer 'purchaseac  'code for sales a/c
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
    DTPicker1_LostFocus

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
    bill1.billcol = 6
    bill1.Billcoltext = ""
    If bill1.billrow < bill1.billrows - 1 Then
    bill1.billrow = bill1.billrow + 1
    End If
Next

Txttotal = ""


bill1.billcol = 2
bill1.billrow = 1
bill1.postext1

DBCombo1.SetFocus

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub DBCombo1_LostFocus()
DBCombo2.BoundText = DBCombo1.BoundText
End Sub

Private Sub DTPicker1_LostFocus()
Dim d As database, r As Recordset, r1 As Recordset, sql As String
Dim total As Single, total1 As Single

sql = "select * from sales where item='" + DBCombo1.BoundText + "' and "
sql = sql + " sales.date=#" & DTPicker1.Month & "/" & DTPicker1.day & "/" & DTPicker1.year & "#"
Set d = OpenDatabase(DBPATH)
Set r1 = d.OpenRecordset("accounts")
r1.Index = "ac_code"

Set r = d.OpenRecordset(sql)

bill1.billrows = 1
bill1.billrows = 2
bill1.billrow = 1

DTPicker2 = DTPicker1
DBCombo2 = DBCombo1

If Not r.NoMatch Then
    Do While Not r.EOF
        bill1.billcol = 0
        bill1.Billcoltext = r!voucher
        bill1.billcol = 1
        bill1.Billcoltext = r!customer
        
        bill1.billcol = 2
        
        r1.Seek "=", r!customer
        If Not r1.NoMatch Then
            bill1.Billcoltext = r1!ac_name
        End If
        bill1.billcol = 3
        bill1.Billcoltext = Format(r!qnty, "0.000")
        total1 = total1 + Val(bill1.Billcoltext)
        bill1.billcol = 4
        bill1.Billcoltext = Format(r!rate, "0.00")
        bill1.billcol = 5
        bill1.Billcoltext = Format(r!rate * r!qnty, "0.00")
        total = total + Val(bill1.Billcoltext)
        
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

Txttotal = Format(total, "0.00")
Txttotal1 = Format(total1, "0.000")

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
bill1.billcols = 7
bill1.billcolwidth = "1000;10;6000;1500;1500;1500;1000;"
bill1.billcolcaption = "||Particulars|Qty|Rate|Amount|"
bill1.database = DBPATH
bill1.table = "select * from accounts where ac_type='K' order by ac_name"
bill1.billlistcolumn = 2
bill1.billlistfield = "ac_name"
bill1.billboundfield = "ac_code"

bill1.billcol = 2
bill1.billleftcol = 0
bill1.billtextpos = 100

Txttotal = ""
DTPicker1 = TRANSACTION_DATE
DTPicker2 = TRANSACTION_DATE

Txtfrom = ""
Txtto = ""
'Text6 = ""
'Text7 = ""
End Sub

Private Sub SSCommand1_Click()

End Sub

Private Sub SSCommand2_Click()
End Sub

Private Sub SSCommand3_Click()

End Sub

Private Sub Txtfrom_LostFocus()
Dim d As database, r As Recordset, r1 As Recordset, sql As String
Dim total As Single, total1 As Single

sql = "select * from sales where item='" + DBCombo1.BoundText + "' and "
sql = sql + " sales.date=#" & DTPicker1.Month & "/" & DTPicker1.day & "/" & DTPicker1.year & "# and sales.voucher>=" & Val(Txtfrom)
Set d = OpenDatabase(DBPATH)
Set r1 = d.OpenRecordset("accounts")
r1.Index = "ac_code"

If Txtfrom = "" Then
    Exit Sub
End If

Set r = d.OpenRecordset(sql)

bill1.billrows = 1
bill1.billrows = 2
bill1.billrow = 1

DTPicker2 = DTPicker1
DBCombo2 = DBCombo1

If Not r.NoMatch Then
    Do While Not r.EOF
        bill1.billcol = 0
        bill1.Billcoltext = r!voucher
        bill1.billcol = 1
        bill1.Billcoltext = r!customer
        
        bill1.billcol = 2
        
        r1.Seek "=", r!customer
        If Not r1.NoMatch Then
            bill1.Billcoltext = r1!ac_name
        End If
        bill1.billcol = 3
        bill1.Billcoltext = Format(r!qnty, "0.000")
        total1 = total1 + Val(bill1.Billcoltext)
        bill1.billcol = 4
        bill1.Billcoltext = Format(r!rate, "0.00")
        bill1.billcol = 5
        bill1.Billcoltext = Format(r!rate * r!qnty, "0.00")
        total = total + Val(bill1.Billcoltext)
        
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

Txttotal = Format(total, "0.00")
Txttotal1 = Format(total1, "0.000")

End Sub

Private Sub Txtto_LostFocus()
Dim d As database, r As Recordset, r1 As Recordset, sql As String
Dim total As Single, total1 As Single

If Txtto <> "" Then
If Val(Txtfrom) > Val(Txtto) Then
    MsgBox " From should be below to", vbCritical
    Exit Sub
End If
End If


sql = "select * from sales where item='" + DBCombo1.BoundText + "' and "
sql = sql + " sales.date=#" & DTPicker1.Month & "/" & DTPicker1.day & "/" & DTPicker1.year & "# and sales.voucher>=" & Val(Txtfrom) & " and sales.voucher<=" & Val(Txtto)
Set d = OpenDatabase(DBPATH)
Set r1 = d.OpenRecordset("accounts")
r1.Index = "ac_code"

Set r = d.OpenRecordset(sql)

bill1.billrows = 1
bill1.billrows = 2
bill1.billrow = 1

DTPicker2 = DTPicker1
DBCombo2 = DBCombo1

If Not r.NoMatch Then
    Do While Not r.EOF
        bill1.billcol = 0
        bill1.Billcoltext = r!voucher
        bill1.billcol = 1
        bill1.Billcoltext = r!customer
        
        bill1.billcol = 2
        
        r1.Seek "=", r!customer
        If Not r1.NoMatch Then
            bill1.Billcoltext = r1!ac_name
        End If
        bill1.billcol = 3
        bill1.Billcoltext = Format(r!qnty, "0.000")
        total1 = total1 + Val(bill1.Billcoltext)
        bill1.billcol = 4
        bill1.Billcoltext = Format(r!rate, "0.00")
        bill1.billcol = 5
        bill1.Billcoltext = Format(r!rate * r!qnty, "0.00")
        total = total + Val(bill1.Billcoltext)
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

Txttotal = Format(total, "0.000")
Txttotal1 = Format(total1, "0.00")
End Sub
