Attribute VB_Name = "Module6"
'splitser by symbol

'***************************************************************
' Name: My own splits function for Vb5
'
' Description: It's my version of the 'splits' Vb 6
'              function for Vb5 Users ...
'
' By: Jérémy cluzel
'     jcluzel@hotmail.com
'
' Inputs:   chaine: string to separate
'           separ: string used as separator, may be more than
'                  one charactere long. Ex: '|', ';', or 'XX'...
'           tableau() : array of string used to store elements
'                       after separation
'           nb_elem : integer used to store the number
'                     of elements (in the array above...)
'
' Returns:  0, if everything works...
'           the number of the error generated else ...
'
'Assumes:None
'
'Side Effects: If you want to use it under Vb6 rename it
'              to 'splits_' or whatever you want...
'
'***************************************************************

Option Explicit
Public Function splits(chaine As String, separ As String, tableau() As String, nb_elem As Integer) As Integer
    On Error GoTo erreur
    Dim pos_act As Integer, pos_occur As Integer
    If Right(chaine, 1) <> separ Then chaine = chaine & separ
    Do
        pos_act = pos_occur + Len(separ)
        pos_occur = InStr(pos_act, chaine, separ)
        If pos_occur <> 0 Then
            ReDim Preserve tableau(nb_elem)
            tableau(nb_elem) = Mid(chaine, pos_act, pos_occur - pos_act)
            nb_elem = nb_elem + 1
        End If
    Loop Until pos_occur = 0
    splits = 0
Exit Function

erreur:
    splits = Err.Number
End Function

