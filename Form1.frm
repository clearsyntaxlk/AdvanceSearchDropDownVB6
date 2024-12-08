VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6195
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11790
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   11790
   StartUpPosition =   3  'Windows Default
   Begin Project1.UserControl1 UserControl11 
      Height          =   2535
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   8415
      _extentx        =   14843
      _extenty        =   4471
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim suggestions() As String
Private Sub Form_Load()
Dim i As Integer
Dim list_ As String
Dim intDimension As Integer
intDimension = 0
For i = 0 To 200

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
    
    
    
    ' Example suggestions (replace with your actual data source)
    'suggestions = Array("Apple", "Banana", "Cherry", "Date", "Grape", "Mango", "Orange", "Peach", "Pear")

    ' Clear the ListBox
    lstSuggestions.Clear
    
    ' Populate ListBox with matching suggestions
    For i = LBound(suggestions) To UBound(suggestions)
        suggestion = suggestions(i)
        If UCase(Left(suggestion, Len(txtInput.Text))) = UCase(txtInput.Text) Then
            lstSuggestions.AddItem suggestion
        End If
    Next i
    
    ' Show or hide the ListBox based on the number of suggestions
    If lstSuggestions.ListCount > 0 Then
        lstSuggestions.Visible = True
    Else
        lstSuggestions.Visible = False
    End If
End Sub


