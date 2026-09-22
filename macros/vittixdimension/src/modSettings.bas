Attribute VB_Name = "modSettings"
Option Explicit
Rem modSettings.bas - centralized settings storage and persistence helpers

Public Type VDT_Settings
    unit As VDTUnit
    decimals As Integer
    Position As String
    fontSize As Integer
    TextWidthPercent As Double
    Gap As Double
    Padding As Double
    CornerRadius As Double
    ShowWidth As Boolean
    ShowHeight As Boolean
    ShowArea As Boolean
    ShowPerimeter As Boolean
    ShowObjectCount As Boolean
    BackgroundBox As Boolean
    RoundedBackground As Boolean
    CreateLayer As Boolean
    RememberSettings As Boolean
    styleName As String
    Template As String
End Type

Public gSettings As VDT_Settings

Public Sub LoadSettings()
    LoadDefaultSettings gSettings
    LoadPersistedSettings gSettings
End Sub

Public Sub SaveSettings()
    If Not gSettings.RememberSettings Then Exit Sub
    PersistSettings gSettings
End Sub

Public Sub LoadDefaultSettings(ByRef settings As VDT_Settings)
    settings.unit = UNIT_MM
    settings.decimals = 2
    settings.Position = "Auto"
    settings.fontSize = 14
    settings.TextWidthPercent = 50
    settings.Gap = 2
    settings.Padding = 2
    settings.CornerRadius = 2
    settings.ShowWidth = True
    settings.ShowHeight = True
    settings.ShowArea = False
    settings.ShowPerimeter = False
    settings.ShowObjectCount = False
    settings.BackgroundBox = True
    settings.RoundedBackground = False
    settings.CreateLayer = True
    settings.RememberSettings = True
    settings.styleName = "Minimal"
    settings.Template = "Size : {W} (W) {U} x {H} (H) {U}"
End Sub

Public Function Settings_PreserveDocumentUnit(doc As Document) As Long
    If doc Is Nothing Then Exit Function
    Settings_PreserveDocumentUnit = doc.unit
End Function

Public Sub Settings_RestoreDocumentUnit(doc As Document, previousUnit As Long)
    On Error Resume Next
    If Not doc Is Nothing Then doc.unit = previousUnit
End Sub

Private Sub PersistSettings(ByRef settings As VDT_Settings)
    
    SaveSetting "Vittix", "DimensionTools", "Unit", CStr(settings.unit)
    SaveSetting "Vittix", "DimensionTools", "Decimals", CStr(settings.decimals)
    SaveSetting "Vittix", "DimensionTools", "Position", settings.Position
    SaveSetting "Vittix", "DimensionTools", "FontSize", CStr(settings.fontSize)
    SaveSetting "Vittix", "DimensionTools", "TextWidthPercent", CStr(settings.TextWidthPercent)
    SaveSetting "Vittix", "DimensionTools", "Gap", CStr(settings.Gap)
    SaveSetting "Vittix", "DimensionTools", "Padding", CStr(settings.Padding)
    SaveSetting "Vittix", "DimensionTools", "CornerRadius", CStr(settings.CornerRadius)
    SaveSetting "Vittix", "DimensionTools", "StyleName", settings.styleName
    SaveSetting "Vittix", "DimensionTools", "Template", settings.Template
    SaveSetting "Vittix", "DimensionTools", "ShowWidth", CStr(settings.ShowWidth)
    SaveSetting "Vittix", "DimensionTools", "ShowHeight", CStr(settings.ShowHeight)
    SaveSetting "Vittix", "DimensionTools", "ShowArea", CStr(settings.ShowArea)
    SaveSetting "Vittix", "DimensionTools", "ShowPerimeter", CStr(settings.ShowPerimeter)
    SaveSetting "Vittix", "DimensionTools", "ShowObjectCount", CStr(settings.ShowObjectCount)
    SaveSetting "Vittix", "DimensionTools", "BackgroundBox", CStr(settings.BackgroundBox)
    SaveSetting "Vittix", "DimensionTools", "RoundedBackground", CStr(settings.RoundedBackground)
    SaveSetting "Vittix", "DimensionTools", "CreateLayer", CStr(settings.CreateLayer)
End Sub

Private Sub LoadPersistedSettings(ByRef settings As VDT_Settings)
    On Error Resume Next
    settings.unit = CLng(GetSetting("Vittix", "DimensionTools", "Unit", CStr(settings.unit)))
    settings.decimals = CInt(GetSetting("Vittix", "DimensionTools", "Decimals", CStr(settings.decimals)))
    settings.Position = GetSetting("Vittix", "DimensionTools", "Position", settings.Position)
    settings.fontSize = CInt(GetSetting("Vittix", "DimensionTools", "FontSize", CStr(settings.fontSize)))
    settings.TextWidthPercent = CDbl(GetSetting("Vittix", "DimensionTools", "TextWidthPercent", CStr(settings.TextWidthPercent)))
    settings.Gap = CDbl(GetSetting("Vittix", "DimensionTools", "Gap", CStr(settings.Gap)))
    settings.Padding = CDbl(GetSetting("Vittix", "DimensionTools", "Padding", CStr(settings.Padding)))
    settings.CornerRadius = CDbl(GetSetting("Vittix", "DimensionTools", "CornerRadius", CStr(settings.CornerRadius)))
    settings.styleName = GetSetting("Vittix", "DimensionTools", "StyleName", settings.styleName)
    settings.Template = GetSetting("Vittix", "DimensionTools", "Template", settings.Template)
    settings.ShowWidth = CBool(GetSetting("Vittix", "DimensionTools", "ShowWidth", CStr(settings.ShowWidth)))
    settings.ShowHeight = CBool(GetSetting("Vittix", "DimensionTools", "ShowHeight", CStr(settings.ShowHeight)))
    settings.ShowArea = CBool(GetSetting("Vittix", "DimensionTools", "ShowArea", CStr(settings.ShowArea)))
    settings.ShowPerimeter = CBool(GetSetting("Vittix", "DimensionTools", "ShowPerimeter", CStr(settings.ShowPerimeter)))
    settings.ShowObjectCount = CBool(GetSetting("Vittix", "DimensionTools", "ShowObjectCount", CStr(settings.ShowObjectCount)))
    settings.BackgroundBox = CBool(GetSetting("Vittix", "DimensionTools", "BackgroundBox", CStr(settings.BackgroundBox)))
    settings.RoundedBackground = CBool(GetSetting("Vittix", "DimensionTools", "RoundedBackground", CStr(settings.RoundedBackground)))
    settings.CreateLayer = CBool(GetSetting("Vittix", "DimensionTools", "CreateLayer", CStr(settings.CreateLayer)))
End Sub

Public Sub ApplySettingsToForm(frm As Object)
    On Error Resume Next
    frm.cmbUnit.value = UnitToText(gSettings.unit)
    frm.cmbDecimals.value = CStr(gSettings.decimals)
    frm.cmbPosition.value = gSettings.Position
    frm.txtFontSize.text = CStr(gSettings.fontSize)
    If Not frm.txtTextWidthPercent Is Nothing Then frm.txtTextWidthPercent.text = CStr(gSettings.TextWidthPercent)
    frm.txtGap.text = CStr(gSettings.Gap)
    frm.txtPadding.text = CStr(gSettings.Padding)
    frm.txtCornerRadius.text = CStr(gSettings.CornerRadius)
    frm.chkShowWidth.value = gSettings.ShowWidth
    frm.chkShowHeight.value = gSettings.ShowHeight
    frm.chkShowArea.value = gSettings.ShowArea
    frm.chkShowPerimeter.value = gSettings.ShowPerimeter
    frm.chkShowObjectCount.value = gSettings.ShowObjectCount
    frm.chkBackgroundBox.value = gSettings.BackgroundBox
    frm.chkRoundedBackground.value = gSettings.RoundedBackground
    frm.chkCreateLayer.value = gSettings.CreateLayer
    frm.chkRememberSettings.value = gSettings.RememberSettings
    frm.txtTemplate.text = gSettings.Template
End Sub

Public Sub ReadSettingsFromForm(frm As Object, ByRef settings As VDT_Settings)
    On Error Resume Next
    Dim parsedDouble As Double
    settings.unit = TextToUnit(frm.cmbUnit.value)
    settings.decimals = CInt(frm.cmbDecimals.value)
    settings.Position = CStr(frm.cmbPosition.value)
    If TryParseDouble(CStr(frm.txtFontSize.text), parsedDouble) Then settings.fontSize = CInt(parsedDouble)
    If Not frm.txtTextWidthPercent Is Nothing Then
        If TryParseDouble(CStr(frm.txtTextWidthPercent.text), parsedDouble) Then settings.TextWidthPercent = parsedDouble
    End If
    If TryParseDouble(CStr(frm.txtGap.text), parsedDouble) Then settings.Gap = parsedDouble
    If TryParseDouble(CStr(frm.txtPadding.text), parsedDouble) Then settings.Padding = parsedDouble
    If TryParseDouble(CStr(frm.txtCornerRadius.text), parsedDouble) Then settings.CornerRadius = parsedDouble
    settings.ShowWidth = CBool(frm.chkShowWidth.value)
    settings.ShowHeight = CBool(frm.chkShowHeight.value)
    settings.ShowArea = CBool(frm.chkShowArea.value)
    settings.ShowPerimeter = CBool(frm.chkShowPerimeter.value)
    settings.ShowObjectCount = CBool(frm.chkShowObjectCount.value)
    settings.BackgroundBox = CBool(frm.chkBackgroundBox.value)
    settings.RoundedBackground = CBool(frm.chkRoundedBackground.value)
    settings.CreateLayer = CBool(frm.chkCreateLayer.value)
    settings.RememberSettings = CBool(frm.chkRememberSettings.value)
    settings.Template = CStr(frm.txtTemplate.text)
End Sub
