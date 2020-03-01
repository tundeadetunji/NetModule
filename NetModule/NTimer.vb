Imports NModule.Methods
Imports Feedback.Feedback
Imports System.Windows.Forms
Public Class NTimer
#Region "Declarations"
	Private f As New Feedback.Feedback
	Public timer_ As Timer, message_str As String, r As Boolean, rCount As Integer, maxRepeatCount As Integer, repeatIndefinately_ As Boolean, rTimer As Timer, rFile As Boolean, feedbackAsFile As Integer, HideApp As Boolean, RepeatFile As Boolean
	Public t_ As Boolean = False

#End Region

	Private Sub SetFiles()
		f.setting_use_voice_feedback = True
		f.selected_language = "English"
	End Sub
	''' <summary>
	''' Builds a reminder timer. Same as BuildTimer.
	''' </summary>
	''' <param name="message_str_"></param>
	''' <param name="timer__"></param>
	''' <param name="repeat_timer"></param>
	''' <param name="repeatCount"></param>
	''' <param name="minutes_interval"></param>
	''' <param name="repeat_"></param>
	''' <param name="repeatIndefinitely"></param>
	''' <param name="MessageIsFile"></param>
	''' <param name="HideApp_"></param>
	''' <param name="RepeatFile_"></param>
	Public Shared Sub StartTimer(message_str_ As String, timer__ As Timer, repeat_timer As Timer, Optional repeatCount As Integer = 3, Optional minutes_interval As Double = 0.125, Optional repeat_ As Boolean = False, Optional repeatIndefinitely As Boolean = False, Optional MessageIsFile As Boolean = False, Optional HideApp_ As Boolean = False, Optional RepeatFile_ As Boolean = False)
		BuildTimer(message_str_, timer__, repeat_timer, repeatCount, minutes_interval, repeat_, repeatIndefinitely, MessageIsFile, HideApp_, RepeatFile_)
	End Sub

	''' <summary>
	''' Builds a reminder timer.
	''' </summary>
	''' <param name="message_str_">What to say.</param>
	''' <param name="timer__">Main Timer</param>
	''' <param name="repeat_timer">Repeat Timer</param>
	''' <param name="repeatCount">How many times to say message_str_. Default is 3.</param>
	''' <param name="minutes_interval">Interval before the next time it begins repeating the message. Default is 15 seconds.</param>
	''' <param name="repeat_">Should the message be repeated after minutes_interval?</param>
	''' <param name="repeatIndefinitely">Should minutes_interval be ignored, and just keep repeating the message? Like a real alarm.</param>
	''' <param name="MessageIsFile">The message is a file, not just string, so start/run the file instead of saying it.</param>
	''' <param name="HideApp_">Attempt to hide the application that is running the file when it is called to start/run the file, if message is file.</param>
	''' <param name="RepeatFile_">Start/Run the file again at the next minutes_interval. Default is false.</param>
	''' <example>
	''' BuildTimer(t.Text, Timer1, Timer2, 3, 1, True)
	''' </example>
	Public Shared Sub BuildTimer(message_str_ As String, timer__ As Timer, repeat_timer As Timer, Optional repeatCount As Integer = 3, Optional minutes_interval As Double = 0.125, Optional repeat_ As Boolean = False, Optional repeatIndefinitely As Boolean = False, Optional MessageIsFile As Boolean = False, Optional HideApp_ As Boolean = False, Optional RepeatFile_ As Boolean = False)
		Dim n_ As New NTimer
		n_.message_str = message_str_
		n_.feedbackAsFile = MessageIsFile
		n_.HideApp = HideApp_
		n_.RepeatFile = RepeatFile_
		n_.timer_ = timer__
		n_.r = repeat_
		n_.rTimer = repeat_timer
		n_.timer_.Interval = minutes_interval * 1000 * 60
		n_.rTimer.Interval = 3000
		n_.maxRepeatCount = repeatCount
		n_.repeatIndefinately_ = repeatIndefinitely

		n_.rCount = 0
		n_.t_ = True
		AddHandler n_.timer_.Tick, New EventHandler(AddressOf n_.TimeTimer_Tick)
		AddHandler n_.rTimer.Tick, New EventHandler(AddressOf n_.RepeatTimer_Tick)

		Dim nt As New NTimer
		nt.SetFiles()

		n_.timer_.Enabled = True

		'repeat, like NetTimer
		'		BuildTimer(t.Text, Timer1, Timer2, 3, 1, True)

	End Sub

	''' <summary>
	''' Stops the timer.
	''' </summary>
	''' <param name="_timer"></param>
	''' <param name="_rTimer"></param>
	Public Shared Sub StopTimer(_timer As Timer, _rTimer As Timer)
		Dim n_ As New NTimer
		_timer.Enabled = False
		_rTimer.Enabled = False
		n_.t_ = False
		n_.message_str = ""
		n_.repeatIndefinately_ = False
		n_.maxRepeatCount = 3
		n_.rCount = 0
	End Sub

	Private Sub RepeatTimer_Tick(sender As Object, e As EventArgs)
		If repeatIndefinately_ = True Then GoTo 2
		If rCount = maxRepeatCount Then
			rTimer.Enabled = False
			rCount = 0
			Exit Sub
		End If
2:
		If feedbackAsFile = False Then
			ReturnFeedback(message_str)
		End If

		rCount += 1
	End Sub
	Private Sub TimeTimer_Tick(sender As Object, e As EventArgs)
		FeedbackProgress()
	End Sub

	Private Sub FeedbackProgress()
		If feedbackAsFile = True Then
			Select Case HideApp
				Case True
					StartFile(message_str, ProcessWindowStyle.Hidden, True)
				Case False
					StartFile(message_str)
			End Select
			timer_.Enabled = RepeatFile
		Else
			rTimer.Enabled = True
			timer_.Enabled = r
		End If
	End Sub
End Class
