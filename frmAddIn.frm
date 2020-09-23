VERSION 5.00
Begin VB.Form frmAddIn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SQL Formatter"
   ClientHeight    =   5535
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   9600
   Icon            =   "frmAddIn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSQLSource 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   9
      Top             =   120
      Width           =   7815
   End
   Begin VB.TextBox txtSQLTarget 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   8
      Top             =   3240
      Width           =   7815
   End
   Begin VB.CommandButton cmdCopyToClipboard 
      Caption         =   "Copy to Clipboard"
      Height          =   375
      Left            =   8040
      TabIndex        =   7
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   7815
      Begin VB.OptionButton Option1 
         Caption         =   "Plain"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Standard Formatted"
         Height          =   495
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Standard Comma seperated"
         Height          =   495
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Code formatted "
         Height          =   495
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Code formatted, comma seprated"
         Height          =   495
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&Close"
      Height          =   375
      Left            =   8040
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VBInstance As VBIDE.VBE
Public Connect As Connect

Option Explicit

Private Sub OKButton_Click()
    Connect.Hide
End Sub

Private Sub cmdCopyToClipboard_Click()
    Clipboard.SetText Me.txtSQLTarget
End Sub

Private Function FormatSQL(sSQLSource As String, Optional bSplitComma, Optional bCode, Optional bPlain) As String

Dim sTemp As String
Dim sPlain As String

' First trip all spaces, crlf etc...
    sTemp = LCase(sSQLSource)
    sTemp = Replace(sTemp, vbCrLf, " ")
    sTemp = Replace(sTemp, vbTab, " ")
    Do While InStr(1, sTemp, "  ")
        sTemp = Replace(sTemp, "  ", " ")
    Loop
    sTemp = Replace(sTemp, " , ", ", ")
    sPlain = sTemp

    
' Now format the right way
    sTemp = Replace(sTemp, "select", "SELECT       ")
    sTemp = Replace(sTemp, "from", vbCrLf & "FROM         ")
    sTemp = Replace(sTemp, "inner join", vbCrLf & "INNER JOIN   ")
    sTemp = Replace(sTemp, "right join", vbCrLf & "RIGHT JOIN   ")
    sTemp = Replace(sTemp, "left join", vbCrLf & "LEFT JOIN    ")
    sTemp = Replace(sTemp, " on ", vbCrLf & "     ON       ")
    sTemp = Replace(sTemp, " and ", vbCrLf & "     AND      ")
    sTemp = Replace(sTemp, "where", vbCrLf & "WHERE        ")
    sTemp = Replace(sTemp, "order by", vbCrLf & "ORDER BY     ")
    
    If Not IsMissing(bSplitComma) Then
        If bSplitComma Then
            sTemp = Replace(sTemp, ",", vbCrLf & ",            ")
        End If
    End If
    
    If Not IsMissing(bCode) Then
        If bCode Then
            sTemp = "sSQL = """ & sTemp
            sTemp = Replace(sTemp, vbCrLf, """ & _ " & vbCrLf & "       """)
            sTemp = sTemp & """"
        End If
    End If
    
    If Not IsMissing(bPlain) Then
        If bPlain = True Then
            sTemp = sPlain
        End If
    End If
    
    FormatSQL = sTemp
    
End Function

Private Sub Option1_Click()
    txtSQLTarget = FormatSQL(txtSQLSource, , , True)
End Sub

Private Sub Option2_Click()
    txtSQLTarget = FormatSQL(txtSQLSource)
End Sub

Private Sub Option3_Click()
    txtSQLTarget = FormatSQL(txtSQLSource, True)
End Sub

Private Sub Option4_Click()
    txtSQLTarget = FormatSQL(txtSQLSource, , True)
End Sub

Private Sub Option5_Click()
        txtSQLTarget = FormatSQL(txtSQLSource, True, True)
End Sub

