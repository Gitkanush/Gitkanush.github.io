' Macro para optimizar macros en cualquier codigo VBA, tomada del libro 
' Learning Excel-VBA

Sub OptimizeVBA(isOn As Boolean)
	Application.Calculation = IIf(isOn, xlCalculationManual, xlCalculationAutomatic)
	Application.EnableEvents = Not(isOn)
	Application.ScreenUpdating = Not(isOn)
	ActiveSheet.DisplayPageBreaks = Not(isOn)
End Sub

' Macro optimizadora para macros grandes
' tomada del libro Learning Excel-VBA

Public Sub FastWB(Optional ByVal opt As Boolean = True)
	With Application
		.Calculation = IIf(opt, xlCalculationManual, xlCalculationAutomatic)
		If .DisplayAlerts <> Not opt Then .DisplayAlerts = Not opt
		If .DisplayStatusBar <> Not opt Then .DisplayStatusBar = Not opt
		If .EnableAnimations <> Not opt Then .EnableAnimations = Not opt
		If .EnableEvents <> Not opt Then .EnableEvents = Not opt
		If .ScreenUpdating <> Not opt Then .ScreenUpdating = Not opt
	End With
	FastWS , opt
End Sub

' Macro optimizadora para macros grandes
' tomada del libro Learning Excel-VBA

Public Sub FastWS(Optional ByVal ws As Worksheet, Optional ByVal opt As Boolean = True)
	If ws Is Nothing Then
		For Each ws In Application.ThisWorkbook.Sheets
			OptimiseWS ws, opt
		Next
	Else
		OptimiseWS ws, opt
	End If
End Sub

' Macro optimizadora para macros grandes
' tomada del libro Learning Excel-VBA

Private Sub OptimiseWS(ByVal ws As Worksheet, ByVal opt As Boolean)
	With ws
		.DisplayPageBreaks = False
		.EnableCalculation = Not opt
		.EnableFormatConditionsCalculation = Not opt
		.EnablePivotTable = Not opt
	End With
End Sub

' Macro para restaurar las funciones de Excel a la normalidad
' tomada del libro Learning Excel-VBA

Public Sub XlResetSettings() 'default Excel settings
	With Application
		.Calculation = xlCalculationAutomatic
		.DisplayAlerts = True
		.DisplayStatusBar = True
		.EnableAnimations = False
		.EnableEvents = True
		.ScreenUpdating = True
		Dim sh As Worksheet
		For Each sh In Application.ThisWorkbook.Sheets
			With sh
				.DisplayPageBreaks = False
				.EnableCalculation = True
				.EnableFormatConditionsCalculation = True
				.EnablePivotTable = True
			End With
		Next
	End With
End Sub