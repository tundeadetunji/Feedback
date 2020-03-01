Imports System.Web.UI.WebControls
Public Class FeedbackWeb
	Private f As New Feedback
	''' <summary>
	''' Lists voice profiles on the PC (Windows), binds it to Control. For FormsControl, use same function in Feedback.
	''' </summary>
	''' <param name="c">Listbox or Combobox</param>
	''' <param name="f_">filter, should it return English only</param>
	''' <returns>Array</returns>
	''' <example>
	''' Dim value_ As Array = f.ListVoices(ComboBox1) 'or ListBox
	''' Dim count_ As Integer = value_(0)
	''' Dim list_ As List(Of String) = value_(1)
	''' </example>

	Public Function ListVoices(Optional c As WebControl = Nothing, Optional f_ As Boolean = True) As Array
		Dim l_ As New List(Of String)()
		With f.voiceCombo
			For i As Integer = 0 To .GetVoices.Count - 1
				If f_ = True Then
					If .GetVoices.Item(i).GetDescription().ToString.Contains("English") Then l_.Add(.GetVoices.Item(i).GetDescription().ToString)
				Else
					l_.Add(.GetVoices.Item(i).GetDescription().ToString)
				End If
			Next
		End With

		If c IsNot Nothing Then
			If TypeOf c Is ListBox Then
				Dim l__ As ListBox = c
				Try
					With l__
						.Items.Clear()
						For li As Integer = 0 To l_.Count - 1
							l__.Items.Add(l_(li).ToString)
						Next
					End With
				Catch ex As Exception
				End Try
			End If
			If TypeOf c Is DropDownList Then
				Dim c__ As DropDownList = c
				Try
					With c__
						.Items.Clear()
						For li As Integer = 0 To l_.Count - 1
							c__.Items.Add(l_(li).ToString)
						Next
					End With
				Catch ex As Exception
				End Try
			End If
		End If

		ListVoices = {f.voiceCombo.GetVoices.Count, l_}

		'		Dim value_ As Array = f.ListVoices(ComboBox1) 'or ListBox
		'		Dim count_ As Integer = value_(0)
		'		Dim list_ As List(Of String) = value_(1)
	End Function

End Class
