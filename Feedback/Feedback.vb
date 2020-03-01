Imports NModule.Methods
Imports SpeechLib
Imports System.Windows.Forms
Public Class Feedback
#Region "Declarations"
	Public WithEvents voiceCombo As New SpVoice
	Public string__ As String = ""
	Public e_ As String = ""

#End Region

#Region "File Paths Variables"

	Private setting_use_voice_feedback_f As Boolean ' = False
	Private setting_accompany_prompt_with_voice_feedback_f As Boolean '= True
	Private setting_use_prompt_f As Boolean ' = False
	Private selected_language_f As String '= "English"

#End Region

#Region "File Paths Properties"
	'speech
	Public Property setting_use_voice_feedback As Boolean
		Get
			Return setting_use_voice_feedback_f
		End Get
		Set(value As Boolean)
			setting_use_voice_feedback_f = value

		End Set
	End Property

	Public Property setting_accompany_prompt_with_voice_feedback As Boolean
		Get
			Return setting_accompany_prompt_with_voice_feedback_f
		End Get
		Set(value As Boolean)
			setting_accompany_prompt_with_voice_feedback_f = value
		End Set
	End Property

	Public Property setting_use_prompt As Boolean
		Get
			Return setting_use_prompt_f
		End Get
		Set(value As Boolean)
			setting_use_prompt_f = value

		End Set
	End Property

	Public Property selected_language As String
		Get
			Return selected_language_f
		End Get
		Set(value As String)
			selected_language_f = value
		End Set
	End Property

#End Region

#Region ""
	Public Sub Welcome(Optional str As String = "Welcome", Optional sound_ As String = Nothing)
		If sound_ IsNot Nothing Then
			Try
				My.Computer.Audio.Play(sound_)
			Catch ex As Exception
			End Try
		End If

		If str.Length > 0 Then fFeedback(str)
	End Sub

	Public Sub Bye(Optional out_timer As Timer = Nothing, Optional str As String = "Bye for now", Optional sound_ As String = Nothing, Optional d As Form = Nothing, Optional close_environment As Boolean = False)

		If sound_ IsNot Nothing Then
			Try
				My.Computer.Audio.Play(sound_)
			Catch ex As Exception
			End Try
		End If

		If str.Length > 0 Then fFeedback(str)

		If out_timer IsNot Nothing Then
			out_timer.Enabled = True
			Exit Sub
		End If

		If close_environment = True Then
			'			g.ExitApp()
			Environment.Exit(0)
			Exit Sub
		End If

		If d IsNot Nothing Then
			d.Close()
			Exit Sub
		End If
	End Sub

#End Region

#Region "Speech"
	''' <summary>
	''' Lists voice profiles on the PC (Windows), binds it to Control. For WebControl, use same function in FeedbackWeb.
	''' </summary>
	''' <param name="c">Listbox or Combobox</param>
	''' <param name="f_">filter, should it return English only</param>
	''' <returns>Array</returns>
	''' <example>
	''' Dim value_ As Array = f.ListVoices(ComboBox1) 'or ListBox
	''' Dim count_ As Integer = value_(0)
	''' Dim list_ As List(Of String) = value_(1)
	''' </example>
	Public Function ListVoices(Optional c As Control = Nothing, Optional f_ As Boolean = True) As Array
		Dim l_ As New List(Of String)()
		With voiceCombo
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
			If TypeOf c Is ComboBox Then
				Dim c__ As ComboBox = c
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

		ListVoices = {voiceCombo.GetVoices.Count, l_}

		'		Dim value_ As Array = f.ListVoices(ComboBox1) 'or ListBox
		'		Dim count_ As Integer = value_(0)
		'		Dim list_ As List(Of String) = value_(1)
	End Function

	Public Sub Say(ByVal textToRead As String, Optional v As Integer = 0, Optional r As Double = 0.2)
		If textToRead.Length < 1 Then Exit Sub

		On Error Resume Next
		If r.ToString.Length < 1 Then r = 0.2

		If v.ToString.Length < 1 Then v = 0

		With voiceCombo
			.Rate = r ' (2 ^ 2 / 20)
			.Voice = .GetVoices.Item(v)

			voiceCombo.Speak(textToRead, SpeechVoiceSpeakFlags.SVSFPurgeBeforeSpeak)

		End With
	End Sub

	Public Sub PauseSay()
		Try
			voiceCombo.Pause()
		Catch
		End Try
	End Sub

	Public Sub ResumeSay()
		'On Error Resume Next
		Try
			voiceCombo.Resume()
		Catch
		End Try
	End Sub


#End Region

#Region "Conversion"
	Public Function ToLanguage(str_ As String) As String
		Return str_

	End Function

#End Region

#Region "Main"
	''' <summary>
	''' Voice Feedback. Will check setting prior. Won't use thread though.
	''' </summary>
	''' <param name="message_"></param>
	Public Shared Function xFeedback(message_ As String) As Boolean
		Dim feedback_ As New Feedback()

		feedback_.string__ = message_
		'		If Length < 1 Then Exit Function

		If feedback_.setting_accompany_prompt_with_voice_feedback = True Or feedback_.setting_use_voice_feedback = True Then
			feedback_.Say(message_)
			Return True
		End If
	End Function

	''' <summary>
	''' Voice Feedback. Use if no language is involved. Won't use thread though.
	''' </summary>
	''' <param name="message_"></param>
	Public Shared Sub ReturnFeedback(message_ As String)
		Dim feedback_ As New Feedback()
		feedback_.Say(message_)
	End Sub

	''' <summary>
	''' Voice Feedback. Will not check setting prior.
	''' </summary>
	''' <param name="message_"></param>
	Public Sub fFeedback(message_ As String)
		string__ = message_
		e_ = "abc"
		VoiceFeedback()
	End Sub

	''' <summary>
	''' Voice Feedback. Will check setting prior.
	''' </summary>
	''' <param name="message_"></param>
	''' <returns></returns>
	Public Shared Function EventFeedback(message_ As String) As Boolean
		Dim feedback_ As New Feedback()

		feedback_.string__ = message_
		'		If Length < 1 Then Exit Function

		If feedback_.setting_accompany_prompt_with_voice_feedback = True Or feedback_.setting_use_voice_feedback = True Then
			feedback_.VoiceFeedback()
			Return True
		End If

	End Function

	''' <summary>
	''' Displays message prompt and/or gives voice feedback. (Relies on previous setting.)
	''' </summary>
	''' <param name="message_">String to put in message box</param>
	''' <param name="voice_feedback_string">String to say</param>
	''' <param name="buttons_">Message box buttons</param>
	''' <param name="echo_">Should message prompt be displayed all the same?</param>
	''' <param name="style_">Message box style</param>
	''' <param name="title">Message box title</param>
	''' <returns>Message box and/or Voice feedback</returns>
	Public Shared Function mFeedback(message_ As String, Optional voice_feedback_string As String = "", Optional buttons_ As MsgBoxStyle = MsgBoxStyle.OkOnly, Optional echo_ As Boolean = False, Optional style_ As MsgBoxStyle = MsgBoxStyle.Information, Optional title As String = "")
		Dim f_ As New Feedback()
		'if voice_feedback_string = null, don't give any voice feedback
		'if echo_ = true, message prompt is compulsory
		'if message_.length > 0, give message prompt
		'ignore entirely if it is not supported

		f_.string__ = voice_feedback_string

		If echo_ Then MsgBox(message_, buttons_ + style_, title) : Exit Function

		'		If f_.string__.Length > 0 Then
		f_.VoiceFeedback()
		'		End If

		If message_.Length < 1 Then Return False ' Exit Function

		Return MsgBox(message_, buttons_ + style_, title)
	End Function

	Public Function Feedback(message_ As String, Optional voice_feedback_string As String = "", Optional buttons_ As MsgBoxStyle = MsgBoxStyle.OkOnly, Optional echo_ As Boolean = False, Optional style_ As MsgBoxStyle = MsgBoxStyle.Information, Optional title As String = "")

		'if voice_feedback_string = null, don't give any voice feedback
		'if echo_ = true, message prompt is compulsory
		'if message_.length > 0, give message prompt
		'ignore entirely if it is not supported

		string__ = voice_feedback_string

		If echo_ Then
			MsgBox(ToLanguage(message_), buttons_ + style_, title)
			Return True
			'			Exit Function
		End If

		If setting_use_voice_feedback = True Then
			VoiceFeedback()
			Return True ' Exit Function
		End If

		If string__.Length > 0 Then
			If setting_accompany_prompt_with_voice_feedback = True Then
				VoiceFeedback()
			End If
		End If


		If message_.Length < 1 Then Return False ' Exit Function

		Return MsgBox(ToLanguage(message_), buttons_ + style_, title)
	End Function

	''' <summary>
	''' Displays message prompt to user. Use this if no language is involved.
	''' </summary>
	''' <param name="message_"></param>
	''' <param name="buttons_"></param>
	''' <param name="style_"></param>
	''' <param name="title"></param>
	''' <returns></returns>
	Public Shared Function m_Feedback(message_ As String, Optional buttons_ As MsgBoxStyle = MsgBoxStyle.OkOnly, Optional style_ As MsgBoxStyle = MsgBoxStyle.Information, Optional title As String = "")

		If message_.Length < 1 Then Return False ' Exit Function

		Return MsgBox(message_, buttons_ + style_, title)
	End Function

	Public Sub VoiceFeedback()
		Try
			If e_.Length > 0 Then
				'				e_ = ""
				GoTo 2
			End If
		Catch
		End Try
		If selected_language.ToLower <> "english" Then Exit Sub
2:
		Dim thread As New System.Threading.Thread(AddressOf SayRequiredText)
		thread.Start()
	End Sub

	Public Sub SayRequiredText()
		Say(string__)
	End Sub


#End Region

End Class
