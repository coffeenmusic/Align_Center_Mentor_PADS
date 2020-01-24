' Description
' Gui with X & Y buttons that will center in either axis. This will center the selected component in between the selected pins.
' It selects the center of the highest and lowest pins selected even if more than 2 are selected.

Public Const PI = 3.14159265359
Public Const MAXFLT = 2^31 - 1

Dim axis As String
Dim selection(100) As String

Public Function Get_Selected_Pins () As Integer
	cnt = 0
	For Each p In ActiveDocument.Pins
		If p.selected Then
			selection(cnt) = p.Name
			cnt = cnt + 1
		End If
	Next p
	
	Get_Selected_Pins = cnt
End Function

Public Function Move_Selected_Cmp (axis, value)
	For Each cmp In ActiveDocument.Components
		If cmp.selected Then
			If axis = "x" Then
				cmp.Move(value, cmp.PositionY, 0)
			Else
				cmp.Move(cmp.PositionX, value, 0)
			End If			
		End If
	Next cmp
End Function

Sub SetCenter (axis)
	Dim x_min, y_min, x_max, y_max As Single
	
	cnt = Get_Selected_Pins
	
	x_min = MAXFLT
	y_min = MAXFLT
	x_max = -MAXFLT
	y_max = -MAXFLT
	
	For i = 0 To cnt - 1
		Set p = ActiveDocument.Pins(selection(i))
		
		If p.PositionX < x_min Then
			x_min = p.PositionX		
		End If
		
		If p.PositionX > x_max Then
			x_max = p.PositionX
		End If
		
		If p.PositionY < y_min Then
			y_min = p.PositionY		
		End If
		
		If p.PositionY > y_max Then
			y_max = p.PositionY
		End If
	
		cnt = cnt + 1
	Next i
	
	x_center = ((x_max - x_min)/2) + x_min
	y_center = ((y_max - y_min)/2) + y_min
	
	If axis = "x" Then
		Move_Selected_Cmp (axis, x_center)
	Else
		Move_Selected_Cmp (axis, y_center)
	End If
	
End Sub

Sub Main
	On Error GoTo EndScript
	Begin Dialog UserDialog 200,84,"Center Component From Pins",.CallbackFunc
		PushButton 40,7,50,21,"X",.btnX 'Center In X
		PushButton 100,7,50,21,"Y",.btnY 'Center In Y
		CancelButton 40,56,110,21
	End Dialog
	Dim dlg As UserDialog
	Dialog dlg
	
	EndScript:
		Exit All
End Sub

Private Function CallbackFunc(DlgItem$, Action%, SuppValue%) As Boolean
	Select Case Action%
		Case 1 ' Dialog box initialization
			' No initialization
		Case 2 ' Value changing or button pressed
			If (DlgItem$ <> "Cancel") Then
				CallbackFunc = True ' Prevent button press from closing the dialog box
				If DlgItem$ = "btnX" Then
					SetCenter "x"
				ElseIf DlgItem$ = "btnY" Then
					SetCenter "y"
				End If
			ElseIf DlgItem$ = "Cancel" Then
			 	CallbackFunc = False ' closing the dialog box
			End If
		Case 3 ' TextBox or ComboBox text changed
		Case 4 ' Focus changed
		Case 5 ' Idle
			Rem CallbackFunc = True ' Continue getting idle actions
	End Select
End Function
