VERSION 5.00
Begin VB.UserControl UserControl1 
   ClientHeight    =   2640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4770
   ScaleHeight     =   2640
   ScaleWidth      =   4770
   Begin VB.CommandButton cmdDropDown 
      Caption         =   "v"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   0
      Width           =   375
   End
   Begin VB.TextBox txtInput 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4335
   End
   Begin VB.ListBox lstSuggestions 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   4695
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    Dim suggestions() As String

Private Sub cmdDropDown_Click()
If UCase(cmdDropDown.Caption) = "V" Then
    lstSuggestions.Visible = True
    cmdDropDown.Caption = "x"
Else
    lstSuggestions.Visible = False
    cmdDropDown.Caption = "v"
End If
End Sub

Private Sub UserControl_Initialize()



Call setComponentWidth

Dim i As Integer
Dim list_ As String
Dim intDimension As Integer
intDimension = 0
For i = 0 To 2000

  ReDim Preserve suggestions(intDimension)
    suggestions(intDimension) = i & " Nex Note"
    intDimension = intDimension + 1
  If Left(CStr(i), 1) = "1" Then
  ReDim Preserve suggestions(intDimension)
      suggestions(intDimension) = "Last of su " & i & "Nex Note"
      intDimension = intDimension + 1
    End If

'    list_ = Arry("Apple", "Banana", "Cherry", "Date", "Grape", "Mango", "Orange", "Peach", "Pear")
Next

For q = 0 To 2000

  ReDim Preserve suggestions(intDimension)
    suggestions(intDimension) = q & " Another List of "
    intDimension = intDimension + 1
  If Left(CStr(q), 1) = "1" Then
  ReDim Preserve suggestions(intDimension)
      suggestions(intDimension) = "Some Kind " & q & "of "
      intDimension = intDimension + 1
    End If

'    list_ = Arry("Apple", "Banana", "Cherry", "Date", "Grape", "Mango", "Orange", "Peach", "Pear")
Next



  ReDim Preserve suggestions(intDimension)
      suggestions(intDimension) = "------Sun----"
      intDimension = intDimension + 1


  ReDim Preserve suggestions(intDimension)
      suggestions(intDimension) = "IN2300075"
      intDimension = intDimension + 1


'--- this is for db recordet
'Dim myIntArray() As Integer
'Dim intDimension As Integer

'intDimension = 0

'Do While Not rstSearchResult.EOF

 'If rstSearchResult(ID) = blah Then
 ' 'Add this Id rstSearchResult(ID) to Array
 ' ReDim Preserve myIntArray(intDimension)
 ' myIntArray(intDimension) = rstSearchResult(ID)
 ' intDimension = intDimension + 1
 'End If

 'Call rstSearchResult.MoveNext
'Loop
'----------------

End Sub


Private Sub setComponentWidth()
    txtInput.Width = UserControl.Width - (cmdDropDown.Width)
    lstSuggestions.Width = UserControl.Width
    cmdDropDown.Left = UserControl.Width - cmdDropDown.Width
    lstSuggestions.Height = UserControl.Height - txtInput.Height
End Sub
Private Sub UserControl_Resize()
setComponentWidth

End Sub



Private Sub lstSuggestions_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Populate the TextBox with the selected suggestion
    txtInput.Text = lstSuggestions.List(lstSuggestions.ListIndex)
    txtInput.SelStart = Len(txtInput.Text) ' Move cursor to the end
    lstSuggestions.Visible = False
End Sub

Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    ' Handle up/down arrow keys to navigate suggestions
    If lstSuggestions.Visible Then
        If KeyCode = vbKeyDown Then
        'If lstSuggestions.ListIndex = -1 Then
        If lstSuggestions.ListCount > 0 Then
            lstSuggestions.SetFocus
            lstSuggestions.ListIndex = lstSuggestions.ListIndex + 1
        'ElseIf lstSuggestions.ListIndex < lstSuggestions.ListCount - 1 Then
               ' lstSuggestions.ListIndex = lstSuggestions.ListIndex + 1
        End If
            KeyCode = 0
        ElseIf KeyCode = vbKeyUp Then
            If lstSuggestions.ListIndex > 0 Then
                lstSuggestions.ListIndex = lstSuggestions.ListIndex - 1
            End If
            KeyCode = 0
        ElseIf KeyCode = vbKeyReturn Then
            ' Select suggestion with Enter key
            If lstSuggestions.ListIndex >= 0 Then
                txtInput.Text = lstSuggestions.List(lstSuggestions.ListIndex)
                txtInput.SelStart = Len(txtInput.Text)
                lstSuggestions.Visible = False
            End If
            KeyCode = 0
        
        End If
    End If
    
    

            If KeyCode = 27 Then  '------ if Esc ky
            txtInput.Text = ""
            lstSuggestions.Visible = False
            'KeyCode = 0
            End If

    
    
End Sub
Private Sub lstSuggestions_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyReturn Then
            ' Select suggestion with Enter key
            If lstSuggestions.ListIndex >= 0 Then
                txtInput.Text = lstSuggestions.List(lstSuggestions.ListIndex)
                txtInput.SelStart = Len(txtInput.Text)
                lstSuggestions.Visible = False
            End If
            KeyCode = 0
        
        ElseIf KeyCode = 27 Then
            txtInput.Text = ""
            lstSuggestions.Visible = False
            KeyCode = 0

        ElseIf KeyCode <> vbKeyDown And KeyCode <> vbKeyUp Then
            If IsCapsLockOn Then
                If Shift = 0 Then
                    ascii = KeyCode
                ElseIf Shift = 1 Then
                    ascii = KeyCode + 32
                End If
            Else
                If Shift = 0 Then
                    ascii = KeyCode + 32
                ElseIf Shift = 1 Then
                    ascii = KeyCode
                End If
            End If
            'txtInput.Text = txtInput.Text & Chr(ascii)
            'txtInput.SetFocus
            'txtInput_KeyDown KeyCode, Shift
            'KeyCode = 0
        End If

End Sub
Private Sub lstSuggestions_KeyPress(KeyAscii As Integer)
'MsgBox (Chr(KeyAscii))
If Trim(txtInput.Text) = "" Then
    Exit Sub
End If
If KeyAscii = 8 Then '------ backspace key press
    txtInput.Text = Left(txtInput.Text, Len(txtInput.Text) - 1)
    txtInput.SetFocus
    txtInput.SelStart = Len(txtInput.Text)
Else
    txtInput.Text = txtInput.Text & (Chr(KeyAscii))
    txtInput.SetFocus
    txtInput.SelStart = Len(txtInput.Text)
End If
End Sub




Private Sub txtInput_Change()

If Trim(txtInput.Text) = "" Then
    Exit Sub
End If

    Dim suggestion As String

    Dim i As Integer
    Dim secondList_()  As String
    Dim intDimension4SecondList As Integer
    intDimension4SecondList = 0
    
    
    ' Example suggestions (replace with your actual data source)
    'suggestions = Array("Apple", "Banana", "Cherry", "Date", "Grape", "Mango", "Orange", "Peach", "Pear")

    ' Clear the ListBox
    lstSuggestions.Clear
    
    ' Populate ListBox with matching suggestions
    If IsArrayEmpty(suggestions) = False Then
    For i = LBound(suggestions) To UBound(suggestions)
    
    If lstSuggestions.ListCount > 10 Then
        Exit For
    End If
    
        suggestion = suggestions(i)
        If UCase(Left(suggestion, Len(txtInput.Text))) = UCase(txtInput.Text) Then
            lstSuggestions.AddItem suggestion
        
        ElseIf Len(txtInput.Text) > 2 Then  '--- add to second list
        Dim splitStr_() As String
        splitStr_ = Split(Trim(txtInput.Text), "%")
        If IsArrayEmpty(splitStr_) = False Then '-------- if it is advacne
            For n = LBound(splitStr_) To UBound(splitStr_)
                If InStr(UCase(suggestion), UCase(splitStr_(n))) > 0 And splitStr_(n) <> "" Then
                    ReDim Preserve secondList_(intDimension4SecondList)
                    secondList_(intDimension4SecondList) = suggestion
                    intDimension4SecondList = intDimension4SecondList + 1
                    Exit For
                End If
            Next n
        Else
            If InStr(UCase(suggestion), UCase(txtInput.Text)) > 0 Then
                ReDim Preserve secondList_(intDimension4SecondList)
                secondList_(intDimension4SecondList) = suggestion
                intDimension4SecondList = intDimension4SecondList + 1
            End If
        End If
        End If
    Next i
    End If
    '------ add second items to list
    If IsArrayEmpty(secondList_) = False Then
    i = 0
    For i = LBound(secondList_) To UBound(secondList_)
        lstSuggestions.AddItem secondList_(i)
    Next i
    End If
'-------- end adding second items to list
    ' Show or hide the ListBox based on the number of suggestions
    If lstSuggestions.ListCount > 0 Then
        lstSuggestions.Visible = True
    Else
        lstSuggestions.Visible = False
    End If
End Sub

Function IsArrayEmpty(arr() As String) As Boolean
    On Error Resume Next
    IsArrayEmpty = (IsEmpty(arr) Or LBound(arr) > UBound(arr))
    If Err.Number <> 0 Then
        IsArrayEmpty = True
    End If
    On Error GoTo 0
End Function

