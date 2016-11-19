' This module contains a few functions and subs to facilitate using hyperlinks to shapes in Visio
' Author: Loïc CORBASSON <loic.corbasson@gmail.com>
' License: MIT

Option Explicit

' Returns the minimum value of an array
Public Function Min(nValues() As Variant)
    Dim n As Variant
    Dim r As Variant
    
    n = nValues(0)
    For Each n In nValues
        If n < r Or IsEmpty(r) Then r = n
    Next n
    
    Min = r
End Function

' Returns the minimum value, except 0, of an array
Public Function MinNot0(nValues() As Variant)
    Dim n As Variant
    Dim r As Variant
    
    For Each n In nValues
        If (n < r Or IsEmpty(r)) And n <> 0 Then r = n
    Next n
    
    MinNot0 = r
End Function

' Returns the maximum value of an array
Public Function Max(nValues() As Variant)
    Dim n As Variant
    Dim r As Variant
    
    n = nValues(0)
    For Each n In nValues
        If n > r Or IsEmpty(r) Then r = n
    Next n
    
    Max = r
End Function

' Renames shapes using the 'title' part of their text: the first line of text, with <<stereotypes>> removed
Public Sub RenameShapesFromText()

    Dim pge As Visio.Page
    Dim shp As Visio.Shape
    Dim shpTitle As String
    Dim nLeftCut As Long, nRightCut As Long
    Dim nLeftCuts() As Variant, nRightCuts() As Variant
    
    For Each pge In Visio.ActiveDocument.Pages
        For Each shp In pge.Shapes
        
            If Not shp.OneD Then
                shpTitle = shp.Text
                shpTitle = Trim(shpTitle)
                nLeftCuts = Array( _
                                InStr(shpTitle, ">>" & Chr(10)) _
                                )
                nLeftCut = Max(nLeftCuts)
                If nLeftCut > 0 Then shpTitle = Mid(shpTitle, nLeftCut + 3)
                shpTitle = Trim(shpTitle)
                nLeftCuts = Array( _
                                InStr(shpTitle, ">>"), _
                                InStr(shpTitle, "»" & Chr(10)) _
                                )
                nLeftCut = Max(nLeftCuts)
                If nLeftCut > 0 Then shpTitle = Mid(shpTitle, nLeftCut + 2)
                shpTitle = Trim(shpTitle)
                nLeftCuts = Array( _
                                InStr(shpTitle, "»") _
                                )
                nLeftCut = Max(nLeftCuts)
                If nLeftCut > 0 Then shpTitle = Mid(shpTitle, nLeftCut + 1)
                shpTitle = Trim(shpTitle)
                nRightCuts = Array( _
                                InStr(shpTitle, ChrW(8232)), _
                                InStr(shpTitle, vbCr), _
                                InStr(shpTitle, vbLf) _
                                )
                nRightCut = MinNot0(nRightCuts)
                If nRightCut > 0 Then shpTitle = Left(shpTitle, nRightCut - 1)
                shpTitle = Trim(shpTitle)
                shp.Name = shpTitle
                shp.NameU = shpTitle
                Debug.Print CStr(shp.ID) + ": " + shp.NameU
            End If
            
'            Debug.Print "OneD " + CStr(shp.OneD) + " : " + shp.Text + " : " + shp.NameID + " : " + shp.NameU + " : " + shp.Name + " : " + CStr(shp.ID)
        Next shp
    Next pge

End Sub
