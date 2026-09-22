Attribute VB_Name = "modLayers"
Option Explicit
Option Private Module
Public Function EnsureDimensionLayer(doc As Document) As Layer
    On Error GoTo ErrHandler
    Dim layerName As String
    Dim lyr As Layer
    layerName = "Dimensions"
    If doc Is Nothing Then Exit Function
    Set lyr = LayerByName(doc, layerName)
    If lyr Is Nothing Then Set lyr = doc.ActivePage.CreateLayer(layerName)
    Set EnsureDimensionLayer = lyr
    Exit Function
ErrHandler:
    LogError "EnsureDimensionLayer", Err.Number, Err.Description
End Function
Private Function LayerByName(doc As Document, layerName As String) As Layer
    Dim lyr As Layer
    Dim candidate As Layer
    On Error Resume Next
    
    ' Try to get it directly first
    Set lyr = doc.ActivePage.Layers(layerName)
    
    ' Fallback to loop if exact match fails
    If lyr Is Nothing Then
        For Each candidate In doc.ActivePage.Layers
            If StrComp(candidate.Name, layerName, vbTextCompare) = 0 Then
                Set lyr = candidate
                Exit For
            End If
        Next candidate
    End If
    
    Set LayerByName = lyr
End Function
