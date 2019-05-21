VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} addForm 
   Caption         =   "Add a New Account"
   ClientHeight    =   10812
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   6084
   OleObjectBlob   =   "addForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "addForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private dataWS As Worksheet, logWS As Worksheet, listWS As Worksheet

Private Sub formAddBtn_Click()
    Dim nextDataRow As Long 'Hold next row to write data to
    Dim nextLogRow As Long 'Hold next row to write log to
    Dim segment As String
    Dim sumSegment As String
    Dim colCounter As Integer
    colCounter = 1
    segment = ""
    
    'Find first row with no content in it within Column A
    nextDataRow = dataWS.Cells(Rows.Count, "A").End(xlUp).row + 1
    nextLogRow = logWS.Cells(Rows.Count, "A").End(xlUp).row + 1

    'Print data to rows
    For Each c In addForm.dataFrame.Controls
        'See if it is a Segment chkbox
        If c.Name = "segmentFrame" Then
            For Each chb In addForm.segmentFrame.Controls
                'Concatenate the chkbox captions into a single cell string
                If chb.Value = "True" Then
                    segment = chb.Caption
                    sumSegment = sumSegment + segment + "," 'Concat
                    dataWS.Cells(nextDataRow, colCounter).Value = sumSegment 'Print
                    logWS.Cells(nextLogRow, colCounter + 1).Value = sumSegment
                End If
            Next chb
            colCounter = colCounter + 1
        'If not the segment section, print into cells
        ElseIf TypeOf c Is MSForms.ComboBox Or TypeOf c Is MSForms.TextBox Or (TypeOf c Is MSForms.CheckBox And c.Top > 320) Then
            dataWS.Cells(nextDataRow, colCounter).Value = c.Value
            logWS.Cells(nextLogRow, colCounter + 1).Value = c.Value
            colCounter = colCounter + 1
        End If
    Next c

    'Add to the drop box lists and sort them if the value doesnt exist already
    FindDupOrAdd companyCBox.Value, companyCBox.Tag
    FindDupOrAdd mspNameCBox.Value, mspNameCBox.Tag
    FindDupOrAdd oemNameCBox.Value, oemNameCBox.Tag

    'Timestamp the log
    logWS.Cells(nextLogRow, "A").Value = Now()
    
    'Prevent data from moving onto new line in cell
    dataWS.Range("A:U").WrapText = False
    logWS.Range("A:V").WrapText = False
    
    'Use to properly close form so that it reinitializes on button press to refresh lists
    Unload Me
    
End Sub


Private Sub InitializeLists()
    Dim counter As Integer
    counter = 1
    Dim tmpList
    Dim lastRow As Long

    'Loop through all comboboxess and add the appropriate list to it
    For Each c In addForm.dataFrame.Controls
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

'Enable chkbox and prep them
Private Sub oemChkBox_Click()
    
    'Call set visible to enable/disable chkboxs
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

Private Sub isgChkBox_Click()
    
    'Call set visible to enable/disable chkboxs
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

'Find if the Account/MSP/OEM/etc added was net new, if so add to appropriate combobox list
Private Sub FindDupOrAdd(searchTerm As String, column As String)

    Dim lastRow
    
    With listWS
        lastRow = .Cells(Rows.Count, column).End(xlUp).row
        
        'Search in the appropriate column for value
        If .Range(column & "2", column & lastRow).Find(searchTerm) Is Nothing Then
            'Add it
            .Range(column & lastRow + 1).Value = searchTerm
            'Sort it A-Z afterwards
            .Range(column & 2, column & (lastRow + 1)).Sort key1:=.Range(column & 1), Order1:=xlAscending
        End If
    End With
    
End Sub







