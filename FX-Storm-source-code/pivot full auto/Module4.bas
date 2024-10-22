Attribute VB_Name = "Module4"
'!!!!!***************!!!!!!!!!******************!!!!!!!!!!!!**********!
'Please read before doing anything with this code

'Disclaimer: This is illegal if excuted on real people and could land you in prison for sure.
'This is intended for educational purposes only. We take no responsibility at all for your actions.
'This code is provided by EEEDS Eagle Eye Digital Security (Oman) for education pupose only.
'For more educational source codes please visit us http://www.digi77.com/training.html
'Dr Jeeni Founder of www.oman0.net & www.digi77.com wishes you good luck :).

'Sharing knowledge is not about giving people something, or getting something from them.
'That is only valid for information sharing.
'Sharing knowledge occurs when people are genuinely interested in helping one another develop new capacities for action;
'it is about creating learning processes.
'Peter Senge
'!!!!!***************!!!!!!!!!******************!!!!!!!!!!!!**********!

'list view  ros color bold changer
Option Explicit

'by John Allan Lee
'Returns
'    1      Modified line
'    0      Did not modify line successfully
'   -1      Internal Function Error
'   -2      Line Index does not exist
Public Function LVW_ModifyLine(lvwListView As ListView, _
                               lngindex As Long, _
                               Optional blnBold As Boolean = False, _
                               Optional strForeColor As String = vbWindowText, _
                               Optional strToolTipText As String = "", _
                               Optional blnErr_ShowFriendly As Boolean, _
                               Optional blnErr_ShowCritical As Boolean) _
                              As Long
On Error GoTo err_LVW_ModifyLine    'initiate error handler
    LVW_ModifyLine = 0              'set default return
    
    'define counter variable
    Dim intColIndex         As Integer
    
    With lvwListView
        'make sure the line exists
        If .ListItems.count < lngindex Then
            LVW_ModifyLine = -2     'set return
            Exit Function
        End If
        'set the first item
        With .ListItems.Item(lngindex)
            .Bold = blnBold
            .ForeColor = strForeColor
            .ToolTipText = strToolTipText
        End With
        'if we don've have children then exit
        If .ColumnHeaders.count < 1 Then
            LVW_ModifyLine = 1      'set positive return
            Exit Function
        End If
        'move through the 'children' of the main item
        For intColIndex = 1 To .ColumnHeaders.count - 1
            'set each child item
            With .ListItems.Item(lngindex).ListSubItems.Item(intColIndex)
                .Bold = blnBold
                .ForeColor = strForeColor
                .ToolTipText = strToolTipText
            End With
        Next intColIndex
    End With
    
    LVW_ModifyLine = 1      'set positive return
    
    Exit Function
err_LVW_ModifyLine:         'error handler
    LVW_ModifyLine = -1     'set internal error return
    'send message to immediate window
    Debug.Print Now & " | Function: & LVW_ModifyLine & | Error: #" & _
                Err.number & vbTab & Err.Description
    'if we want to show critical messages to the user
    If blnErr_ShowCritical = True Then
        'notify the user
        MsgBox "Error: #" & Err.number & vbTab & Err.Description & _
               vbCrLf & vbCrLf & Now, _
               vbOKOnly + vbCritical, _
               "Function: LVW_ModifyLine"
    End If
    Err.Clear   'clear the error object
    
End Function


