VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} editForm 
   Caption         =   "Edit an Account"
   ClientHeight    =   10812
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   6084
   OleObjectBlob   =   "editForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "editForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private dataWS As Worksheet, logWS As Worksheet, listWS As Worksheet
Private foundRow

'Prepopulate data based on selected or searched Account combobox item
Private Sub companyCBox_AfterUpdate()

    Dim lookupName
    Dim foundCell
    Dim colCounter
    Dim segmentSplit() As String
    colCounter = 1
    
    lookupName = companyCBox.Value
   
    'Search company names and return the location of the selected one
    'dataWS.Cells(1, 1).Value = dataWS.Range("A:A").Find(lookupName).Address(False, False) 'False removes the cell absolutes
    foundCell = dataWS.Range("A:A").Find(lookupName).Address(False, False)
    foundRow = dataWS.Range(foundCell).row
    'dataWS.Cells(2, 1).Value = foundRow

    For Each c In editForm.dataFrame.Controls
         'Check if part of Segments data
         If c.Name = "segmentFrame" Then
            'Break the segments string into an array
            segmentSplit = Split(dataWS.Cells(foundRow, colCounter).Value, ",")
            
            'Loop the checkboxes for each segment in the array to populate the form
            For Each chb In editForm.segmentFrame.Controls
                For i = LBound(segmentSplit()) To UBound(segmentSplit())
                    'If a checkbox caption equals a segment, then set checkbox to true
                    If chb.Caption = segmentSplit(i) Then
                    
                        chb.Value = "True"
                        
                    End If
                Next i
            Next chb
            
            colCounter = colCounter + 1
        'If not segment data,just set the values on the row being edited
        ElseIf TypeOf c Is MSForms.ComboBox Or TypeOf c Is MSForms.TextBox Or (TypeOf c Is MSForms.CheckBox And c.Top > 320) Then
    
            c.Value = dataWS.Cells(foundRow, colCounter).Value
    
            colCounter = colCounter + 1
            
        End If
    Next c
End Sub


Private Sub formSaveBtn_Click()

    Dim nextLogRow As Long 'Hold next row to write log to
    Dim segment As String
    Dim sumSegment As String
    Dim colCounter As Integer
    colCounter = 1
    segment = ""

    'Find first row with no content in it within Column A
    nextLogRow = logWS.Cells(Rows.Count, "A").End(xlUp).row + 1
    
    'Print data to rows
    For Each c In editForm.dataFrame.Controls
        'See if it is a Segment chkbox
        If c.Name = "segmentFrame" Then
            For Each chb In editForm.segmentFrame.Controls
                'Concatenate the chkbox captions into a single string
                If chb.Value = "True" Then
                    segment = chb.Caption
                    sumSegment = sumSegment + segment + "," 'Concat
                    dataWS.Cells(foundRow, colCounter).Value = sumSegment 'Print
                    logWS.Cells(nextLogRow, colCounter + 1).Value = sumSegment
                End If
            Next chb
            colCounter = colCounter + 1
        'If not the Segment section, print into cells
        ElseIf TypeOf c Is MSForms.ComboBox Or TypeOf c Is MSForms.TextBox Or (TypeOf c Is MSForms.CheckBox And c.Top > 320) Then
            dataWS.Cells(foundRow, colCounter).Value = c.Value
            logWS.Cells(nextLogRow, colCounter + 1).Value = c.Value
            colCounter = colCounter + 1
        End If
    Next c

    'Timestamp the log
    logWS.Cells(nextLogRow, "A").Value = Now()
    
    dataWS.Range("A:U").WrapText = False
    logWS.Range("A:V").WrapText = False
    
    'Use to properly close form so that it reinitializes on button press to refresh menus
    Unload Me
    
End Sub


Private Sub InitializeLists()
    Dim counter As Integer
    counter = 1
    Dim tmpList
    Dim lastRow As Long

   'Loop through all comboboxess and add the appropriate list to it
   For Each c In editForm.dataFrame.Controls
        If TypeOf c Is MSForms.ComboBox Then 'Loop through CBox controls
            With listWS
                'Find and select the range of the individual list and store it in temp var
                tmpList = .Range(.Cells(2, counter), .Cells(.Cells(Rows.Count, counter).End(xlUp).row, counter)).Value2
            End With
            'Set property
            c.list = tmpList
            counter = counter + 1 'Inc
        End If
    Next c
End Sub

'Enable checkbox and prep it
Private Sub mspChkBox_Click()
    'Call set visible to enable/disable chkbox
    SetVisible mspChkBox, mspNameCBox
    
    If mspChkBox.Value = "True" Then
        mspNameCBox.SetFocus
    Else
        mspNameCBox.Value = ""
    End If

End Sub

'Enable checkbox and prep it
Private Sub oemChkBox_Click()
    'Call set visible to enable/disable chkbox
    SetVisible oemChkBox, oemNameCBox
    SetVisible oemChkBox, oemContactTxt
    SetVisible oemChkBox, oemContactLbl
    
    If oemChkBox.Value = "True" Then
        oemNameCBox.SetFocus
    Else
        oemNameCBox.Value = ""
        oemContactTxt.Value = ""
    End If

End Sub

'Enable checkbox and prep it
Private Sub isgChkBox_Click()
    'Call set visible to enable/disable chkbox
    SetVisible isgChkBox, isgTxt
    SetVisible isgChkBox, bdmLbl
    SetVisible isgChkBox, bdmTxt
    
    If isgChkBox.Value = "True" Then
        isgTxt.SetFocus
    Else
        isgTxt.Value = ""
        bdmTxt.Value = ""
    End If

End Sub

Private Sub UserForm_Initialize()
    
    'Store the 3 worksheets
    Set dataWS = Worksheets("Data")
    Set logWS = Worksheets("Log")
    Set listWS = Worksheets("Lists")
    
    'Assign preset values below
    addForm.startDateTxt.Value = Date
    companyCBox.SetFocus
    
    'Call method to setup combobox data lists
    InitializeLists

End Sub

'Change visibility of a control based on chkbox boolean
Private Sub SetVisible(chk As MSForms.CheckBox, c As Control)

    If chk.Value = "True" Then
        c.Visible = True
    Else
        c.Visible = False
    End If

End Sub
