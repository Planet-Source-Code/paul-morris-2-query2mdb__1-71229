VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   ScaleHeight     =   6945
   ScaleWidth      =   9780
   StartUpPosition =   3  'Windows Default
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex1 
      Height          =   6675
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   11774
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'I needed to query 2 password protected Access databases with a single SQL statement.
'
'The Microsoft web page http://support.microsoft.com/kb/113701 was helpful and
'gave me hope that it may be possible, but not too clear.
'
'I use ADO with VB6 whereas the Microsoft example was for VB3 and the old ODBC.
'The Microsoft example worked almost straightaway for me if the databases had
'no password protection, it was the password protection that made it more
'difficult.
'
'Eventually after quite a bit of experimentation I cracked it and I thought I
'must share this with fellow coders in case they have the same requirement.
'
'
'
'There are 2 databses with this code: -
'1. BookSale_2002.mdb   -  password = ABCD
'2. BS2.mdb             -  password = 1234
'They are the BookSale.mdb database, supplied Microsoft in Visual Studio, with
'the tables split between them.
'The 2 databases should be located in the same folder as the VB files.
'



Private Sub Form_Load()
Dim db As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sSQL As String

Dim sDB2 As String
Dim sDBpath As String

   sDBpath = App.Path

   'connection for the 1st database
   Set db = New ADODB.Connection
   db.CursorLocation = adUseClient
   db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDBpath _
                     & "\BookSale_2002.mdb;Jet OLEDB:Database Password=ABCD"
                     
   
   
   '====================================================================================
   'connection string for 2nd database to queried
   sDB2 = "[;database=" & sDBpath _
                     & "\BS2.mdb;pwd=1234]"
                     
   
   sSQL = "SELECT Title, Author, [Name] AS Publisher, ti.ISBN," _
                  & " Format(Price, '00.00') AS [Price($)]" _
                  & "" _
         & " FROM Authors au, Publishers pu," _
                  & " " & sDB2 & ".Titles ti, " & sDB2 & ".[Title Author] ta" _
                  & "" _
         & " WHERE pu.PubID = ti.PubID" _
                  & " AND ta.ISBN = ti.ISBN" _
                  & " AND au.au_ID = ta.au_ID" _
                  & "" _
         & " ORDER BY Title"


   Set rs = New ADODB.Recordset                             'initialise the recordset
   rs.Open sSQL, db, adOpenStatic, adLockOptimistic         'open the recordset

   
   
   '====================================================================================
   'grid settings
   With Flex1
      .FixedCols = 0
      .ScrollTrack = True
      If rs.RecordCount = 0 Then
         .Rows = 2
      Else
         Set .DataSource = rs
      End If
      
      .ColWidth(0) = 4000                 'Title
      .ColWidth(1) = 1600                 'Author
      .ColWidth(2) = 1200                 'Publisher
      .ColWidth(3) = 1000                 'ISBN
      .ColWidth(4) = 700                  'Price
      
      .AllowUserResizing = flexResizeColumns
   End With
   
   rs.Close
   db.Close
End Sub
