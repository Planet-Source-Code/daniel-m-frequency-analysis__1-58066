VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmFrequency 
   Caption         =   "Frequency Analysis Chart Generator"
   ClientHeight    =   10155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9645
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   10155
   ScaleWidth      =   9645
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cDlgSave 
      Left            =   7920
      Top             =   9600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export Chart"
      Height          =   375
      Left            =   4440
      TabIndex        =   11
      Top             =   9720
      Width           =   1815
   End
   Begin VB.TextBox txtExcludePercent 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Text            =   "!,. -:;"
      Top             =   4800
      Width           =   6375
   End
   Begin VB.ListBox lstChart 
      Columns         =   2
      Height          =   4140
      ItemData        =   "frmFrequency.frx":0000
      Left            =   120
      List            =   "frmFrequency.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   5520
      Width           =   6255
   End
   Begin VB.OptionButton OptContain 
      Caption         =   "Display All Containing Characters"
      Height          =   255
      Index           =   2
      Left            =   6600
      TabIndex        =   4
      Top             =   4440
      Width           =   2895
   End
   Begin VB.OptionButton OptContain 
      Caption         =   "Display AlphaNumeric characters Only"
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   3
      Top             =   4440
      Width           =   3135
   End
   Begin VB.OptionButton OptContain 
      Caption         =   "Display Alpha-characters Only"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   4440
      Value           =   -1  'True
      Width           =   2535
   End
   Begin VB.CommandButton cmdGenChart 
      Caption         =   "Generate Chart"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   4080
      Width           =   2055
   End
   Begin VB.TextBox txtCipher 
      Height          =   4095
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmFrequency.frx":0004
      Top             =   0
      Width           =   9615
   End
   Begin VB.Label lblHeader2 
      Height          =   255
      Left            =   3120
      TabIndex        =   10
      Top             =   5280
      UseMnemonic     =   0   'False
      Width           =   3015
   End
   Begin VB.Label lblHeader 
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   5280
      UseMnemonic     =   0   'False
      Width           =   3015
   End
   Begin VB.Label lblDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmFrequency.frx":2AC5
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   6480
      TabIndex        =   8
      Top             =   5640
      Width           =   3135
   End
   Begin VB.Label lblExclude 
      Caption         =   "Exclude Characters from Percentages:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   4845
      Width           =   3015
   End
End
Attribute VB_Name = "frmFrequency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExport_Click()
If lstChart.ListCount <> 0 Then
    With cDlgSave 'Save file
        .FileName = App.Path & "Frequency Chart"
        .Filter = "Text Document (*.txt)|*.txt|All Files (*.*)|*.*|"
        .ShowSave
    End With

    If vbCancel <> True Then
        Open cDlgSave.FileName For Append As #1
            For i = 0 To lstChart.ListCount - 1
                Print #1, lstChart.List(i) & vbNewLine
            Next i
        Close #1
        
        MsgBox "Successfully Saved!", vbInformation, "Chart Saved"
    End If
End If
End Sub

Private Sub cmdGenChart_Click()
If Len(txtCipher.Text) <> 0 Then
    GenerateFrequencyChart
Else
    MsgBox "Cannot Generate Chart, No Data!", vbInformation, "Cannot Generate"
    txtCipher.SetFocus
End If

End Sub

Private Function GenerateFrequencyChart()
Dim CipherText As String, arrChars() As String, arrFreq() As Long, i As Long, j As Long, lngTotalDenominator As Long
Dim addChar As Boolean, firstInstance As Boolean, chartType As Long, lngSubDenominator As Long

For i = 0 To 2
    If OptContain(i).Value = True Then
        chartType = i
        Exit For
    End If
Next i

CipherText = UCase(txtCipher.Text) 'we only need one case of each letter to decipher text
ReDim arrChars(0) 'redim to hold data
ReDim arrFreq(0)
firstInstance = True
For i = 1 To Len(CipherText)
    addChar = True
    
    For j = 0 To UBound(arrChars)
        If Mid(CipherText, i, 1) = arrChars(j) Then
            arrFreq(j) = arrFreq(j) + 1
            addChar = False
            Exit For
        End If
    Next j
    
    If addChar = True Then
        If firstInstance = True Then
            firstInstance = False
            arrChars(UBound(arrChars)) = Mid(CipherText, i, 1)
            arrFreq(UBound(arrFreq)) = 1
        Else
            ReDim Preserve arrChars(UBound(arrChars) + 1)
            ReDim Preserve arrFreq(UBound(arrFreq) + 1)
            arrChars(UBound(arrChars)) = Mid(CipherText, i, 1)
            arrFreq(UBound(arrFreq)) = 1
        End If
    End If
Next i


For i = 0 To UBound(arrChars)
    Select Case chartType 'this looks similar to the next step but instead of generating the chart, this
        Case 0            'will count the number of characters we do not wish to include in the frequency
            If Asc(arrChars(i)) < 65 Or Asc(arrChars(i)) > 90 Then
                blnProceed = False
            Else
                blnProceed = True
            End If
        
        Case 1
            If Asc(arrChars(i)) < 48 Or Asc(arrChars(i)) > 57 And Asc(arrChars(i)) < 65 Or Asc(arrChars(i)) _
            > 90 Then
                blnProceed = False
            Else
                blnProceed = True
            End If
        
        Case 2
            blnProceed = True
    End Select
    
    If blnProceed = False Then
        lngSubDenominator = lngSubDenominator + arrFreq(i)
    End If
Next i

lngTotalDenominator = Len(CipherText) - lngSubDenominator

For i = 0 To UBound(arrChars)
    Select Case chartType 'handles what is going to be on the frequency chart
        Case 0
            If Asc(arrChars(i)) < 65 Or Asc(arrChars(i)) > 90 Then
                blnProceed = False
            Else
                blnProceed = True
            End If
        
        Case 1
            If Asc(arrChars(i)) < 48 Or Asc(arrChars(i)) > 57 And Asc(arrChars(i)) < 65 Or Asc(arrChars(i)) _
            > 90 Then
                blnProceed = False
            Else
                blnProceed = True
            End If
        
        Case 2
            blnProceed = True
    End Select
    
    If blnProceed = True Then
        lstChart.AddItem arrChars(i) & vbTab & arrFreq(i) & vbTab & Round((arrFreq(i) / lngTotalDenominator) * 100, 2) & "%"
    End If
Next i
End Function


Private Sub Form_Load()
lblHeader.Caption = "Chr" & Space(10) & "Occurs" & Space(5) & "Frequency"
lblHeader2.Caption = "Chr" & Space(10) & "Occurs" & Space(5) & "Frequency"
End Sub

