Imports System.Runtime.InteropServices
Imports System.Globalization
Public Class Form_Calculator ' Created by: Joshua C. Magoliman
    ' This line of code will able to use the Metro UI Form
    Inherits MetroFramework.Forms.MetroForm

#Region "Fields"
    ' All these fields will hold specific values
    Dim firstGivenNumber As Double, secondGivenNumber As Double, answer As Double
    ' This field will hold the arithmetic operator
    Dim arithmeticOperator As Char = " "
    ' This field will hold the value, if the btnDot is clicked or not
    Dim isbtnDotClicked As Boolean = False
    ' Create an object called sound
    Dim sound As New System.Media.SoundPlayer()
    ' Create an object called resourcesManager
    Dim resourcesManager = New Resources.ResourceManager("SIMPLECALCULATOR_IN_VB.Resources", System.Reflection.Assembly.GetExecutingAssembly)
#End Region

#Region "Event Handler Methods (Default Naming Convention)"
    Private Sub Form_Calculator_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Disable this textbox
        txtFirst.Enabled = False
        ' Hide this textbox
        txtSecond.Visible = False
        ' Disable this textbox
        txtSecond.Enabled = False
        ' Use the 'MetroStyleManager1' as a style manager in this form
        Me.StyleManager = MetroStyleManager1
        ' Set the text of this button to "DARK"
        btnChangeTheme.Text = "DARK"
        ' Set the location of sound and the filename of sound
        sound.SoundLocation = Application.StartupPath & "\WAV\BUTTONCLICK.wav"
        ' Invoke the Object called 'sound' and use the Built-in Method called Load()
        sound.Load()
        ' Set the icon of this Form
        Me.Icon = resourcesManager.GetObject("CALCULATOR")
        ' Invoke this API Method called 'MakeThisFormRoundedRectangle' and give all necessary values in arguments
        Region = System.Drawing.Region.FromHrgn(MakeThisFormRoundedRectangle(0, 0, Width, Height, 30, 30))
    End Sub
    Private Sub btnChangeTheme_Click(sender As Object, e As EventArgs) Handles btnChangeTheme.Click
        ' This local variable will hold the value of the button called btnChangeTheme
        Dim valueText As String = btnChangeTheme.Text
        ' Check the value of local variable called valueText
        Select Case valueText
            ' If the value of local variable called valueText is "DARK"
            Case "DARK"
                ' Change the MetroThemeStyle to Dark
                MetroStyleManager1.Theme = MetroFramework.MetroThemeStyle.Dark
                ' Set the text of btnChangeTheme to "LIGHT"
                btnChangeTheme.Text = "LIGHT"

                ' If the value of local variable called valueText is "LIGHT"
            Case "LIGHT"
                ' Change the MetroThemeStyle to Light
                MetroStyleManager1.Theme = MetroFramework.MetroThemeStyle.Light
                ' Set the text of btnChangeTheme to "DARK"
                btnChangeTheme.Text = "DARK"
        End Select
        ' Invoke the Object called 'sound' and use the Built-in Method called Play()
        sound.Play()
    End Sub
    ' This Event Handler Method will handles all buttons have number digits from 1 until 9 and 0 
    Private Sub btnNumbersClicking(sender As Object, e As EventArgs) Handles btnOne.Click, btnTwo.Click, btnThree.Click, btnFour.Click, btnFive.Click, btnSix.Click, btnSeven.Click, btnEight.Click, btnNine.Click, btnZero.Click
        ' If textbox called txtFirst is visible and textbox called txtSecond is not visible
        If txtFirst.Visible = True And txtSecond.Visible = False Then
            ' Add the text value of Parameter Object called 'sender' in textbox called txtFirst
            txtFirst.Text = txtFirst.Text & sender.text

            ' If textbox called txtFirst and textbox called txtSecond are both visible
        ElseIf txtFirst.Visible = True And txtSecond.Visible = True Then
            ' Add the text value of Parameter Object 'sender' in textbox called txtSecond
            txtSecond.Text = txtSecond.Text & sender.text
        End If

        ' Invoke the object called 'sound' and use the Built-in Method called Play()
        sound.Play()
        ' Invoke this User Defined Method called AddCommasFortxtFirst()
        AddCommasFortxtFirst()
        ' Invoke this User Defined Method called AddCommasFortxtSecond()
        AddCommasFortxtSecond()
    End Sub
    Private Sub btnPlus_Click(sender As Object, e As EventArgs) Handles btnPlus.Click
        CheckTheTwoTextBox("+", " +")
    End Sub
    Private Sub btnMinus_Click(sender As Object, e As EventArgs) Handles btnMinus.Click
        CheckTheTwoTextBox("-", " -")
    End Sub
    Private Sub btnTimes_Click(sender As Object, e As EventArgs) Handles btnTimes.Click
        CheckTheTwoTextBox("*", " x")
    End Sub
    Private Sub btnDivide_Click(sender As Object, e As EventArgs) Handles btnDivide.Click
        CheckTheTwoTextBox("/", " /")
    End Sub
    Private Sub btnDot_Click(sender As Object, e As EventArgs) Handles btnDot.Click
        ' If the text of textbox called txtFirst is not empty
        If Not txtFirst.Text = "" Then
            ' If the textbox called txtSecond is not visible and empty
            If txtSecond.Visible = False And txtSecond.Text = "" Then
                ' Check if this textbox called txtFirst have a "."
                If InStr(txtFirst.Text, ".") > 0 Then
                    '  // End the execution of this Event Handler Method (Default Naming Convention) called btnDot_Click
                    Exit Sub

                    ' Check if this textbox called txtFirst don't have a "."
                Else
                    ' Set "." in textbox called txtFirst
                    txtFirst.Text = txtFirst.Text & "."
                End If

                ' If the textbox called txtSecond is visible and not empty
            ElseIf txtSecond.Visible = True And Not txtSecond.Text = "" Then
                ' Check if textbox called txtSecond have a "."
                If InStr(txtSecond.Text, ".") > 0 Then
                    '  // End the execution of this Event Handler Method (Default Naming Convention) called btnDot_Click
                    Exit Sub

                    ' Check if this textbox called txtSecond don't have a "."
                Else
                    ' Set "." in textbox called txtSecond
                    txtSecond.Text = txtSecond.Text & "."
                End If
            End If
            ' Re assign the field called 'isbtnDotClicked' to true
            isbtnDotClicked = True
        End If
        ' Invoke the object called 'sound' and use the Built-in Method called Play()
        sound.Play()
    End Sub
    Private Sub btnPositiveOrNegative_Click(sender As Object, e As EventArgs) Handles btnPositiveOrNegative.Click
        ' If textbox called txtFirst is not empty and textbox called txtSecond is not visible
        If Not txtFirst.Text = "" And txtSecond.Visible = False Then
            ' Set the positive or negative sign in textbox called txtFirst
            txtFirst.Text = -1 * txtFirst.Text

            ' If textbox called txtFirst is not empty and textbox called txtSecond is visible
        ElseIf Not txtFirst.Text = "" And txtSecond.Visible = True Then
            ' Set the positive or negative sign in textbox called txtSecond
            txtSecond.Text = -1 * txtSecond.Text
        End If
        ' Invoke the object called 'sound' and use the Built-in Method called Play()
        sound.Play()
    End Sub
    Private Sub btnErase_Click(sender As Object, e As EventArgs) Handles btnErase.Click
        ' If the textbox called txtSecond is visible
        If txtSecond.Visible = True Then
            ' If the text value of textbox called txtSecond is not empty
            If Not txtSecond.Text = "" Then
                ' Erase 1 value of text in textbox called txtSecond
                txtSecond.Text = txtSecond.Text.Remove(txtSecond.Text.Count - 1)
                ' Invoke the User Defined Method called AddCommasFortxtSecond()
                AddCommasFortxtSecond()
            End If

            ' If the textbox called txtSecond is not visible
        Else
            ' If the text value of textbox called txtFirst is not empty
            If Not txtFirst.Text = "" Then
                ' Erase 1 value of text in textbox called txtFirst
                txtFirst.Text = txtFirst.Text.Remove(txtFirst.Text.Count - 1)
                ' Invoke the User Defined Method called AddCommasFortxtFirst()
                AddCommasFortxtFirst()
            End If
        End If
        ' Invoke the object called 'sound' and use the Built-in Method called Play()
        sound.Play()
    End Sub
    Private Sub btnClearEntry_Click(sender As Object, e As EventArgs) Handles btnClearEntry.Click
        ' If textbox called txtSecond is visible
        If txtSecond.Visible = True Then
            ' Clear the text of this textbox called txtSecond
            txtSecond.Clear()
            ' Set the focus in this textbox called txtSecond
            txtSecond.Focus()

            ' If the textbox called txtSecond is not visible
        Else
            ' Clear the text of this textbox called txtFirst
            txtFirst.Clear()
            ' Set the focus in this textbox called txtSecond
            txtSecond.Clear()
            ' Hide the textbox called txtSecond
            txtSecond.Visible = False
        End If
        ' Invoke the object called 'sound' and use the Built-in Method called Play()
        sound.Play()
    End Sub
    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        Clear()
    End Sub
    Private Sub btnEqual_Click(sender As Object, e As EventArgs) Handles btnEqual.Click
        ' Invoke the User Defined Method called CalculateTheAnswer()
        CalculateTheAnswer()
    End Sub
    Private Sub btnCopyToClipBoard_Click(sender As Object, e As EventArgs) Handles btnCopyToClipBoard.Click
        ' If the text value of textbox called txtFirst is not empty
        If txtFirst.Text <> String.Empty Then
            ' Get the value of text of textbox called txtFirst and set the value of it to the local variable called getTheValue
            Dim getTheValue As String = txtFirst.Text
            ' Declare a new local variable that will hold the new value
            Dim newValue As String

            ' If the value of local variable called getTheValue have contain " +"
            If getTheValue.Contains(" +") Then
                ' Find the value " +" and replace it with ""
                newValue = getTheValue.Replace(" +", "")

                ' If the value of local variable called getTheValue have contain " -"
            ElseIf getTheValue.Contains(" -") Then
                ' Find the value " -" and replace it with ""
                newValue = getTheValue.Replace(" -", "")

                ' If the value of local variable called getTheValue have contain " x"
            ElseIf getTheValue.Contains(" x") Then
                ' Find the value " x" and replace it with ""
                newValue = getTheValue.Replace(" x", "")

                ' If the value of local variable called getTheValue have contain " /"
            ElseIf getTheValue.Contains(" /") Then
                ' Find the value " /" and replace it with ""
                newValue = getTheValue.Replace(" /", "")

                ' If the local variable don't have a value of " +" or " -" or " x" or " /"
            Else
                ' Get the value of local variable called getTheValue and set it in the local variable called newValue
                newValue = getTheValue
            End If

            ' Set the text of value of local variable called newValue using the Built-in Method called SetText
            Clipboard.SetText(newValue)
            ' Show this messagebox
            MsgBox("SUCCESSFULLY COPIED!", MsgBoxStyle.Information, "ATTENTION")

            ' If the text value of textbox called txtFirst is empty 
        Else
            ' Clear the content of Clipboard class using the Built-in Method called Clear
            Clipboard.Clear()
        End If

        ' Set the Focus in button called btnEqual
        btnEqual.Focus()
        ' Invoke the Object called 'sound' and use the Built-in Method called Play()
        sound.Play()
    End Sub
#End Region

#Region "API Method"
    ' This API Method will make this Form Rounded Rectangle. Note: Do not change this ' Alias "CreateRoundRectRgn" '
    Private Declare Function MakeThisFormRoundedRectangle Lib "Gdi32.dll" Alias "CreateRoundRectRgn" (ByVal param_LeftSide As Integer, ByVal param_TopSide As Integer, ByVal param_RightSide As Integer, ByVal param_BottomSide As Integer, ByVal param_WidthSize As Integer, ByVal param_HeightSize As Integer) As IntPtr
#End Region

#Region "User Defined Methods"
    Private Sub AddCommasFortxtFirst()
        ' If the text length of textbox called txtFirst is greater than or equal to 4 and the textbox called txtSecond is not visible
        If txtFirst.Text.Trim.Length >= 4 And txtSecond.Visible = False Then
            ' Declare a local variable for converting the format to data type double
            Dim patternFormat As Double = Convert.ToDouble(txtFirst.Text)
            ' If the textbox called txtFirst don't contains "."
            If Not txtFirst.Text.Contains(".") Then
                ' Convert the string format of the text of textbox called txtFirst 
                txtFirst.Text = patternFormat.ToString("#,##0")
            End If
        End If
    End Sub
    Private Sub AddCommasFortxtSecond()
        ' If the text length of textbox called txtSecond is greater than or equal to 4
        If txtSecond.Text.Trim.Length >= 4 Then
            ' Declare a local variable for converting the format
            Dim patternFormat As Double = Convert.ToDouble(txtSecond.Text)
            ' If the textbox called txtSecond don't contains "."
            If Not txtSecond.Text.Contains(".") Then
                ' Convert the string format of the text of textbox called txtSecond 
                txtSecond.Text = patternFormat.ToString("#,##0")
            End If
        End If
    End Sub
    Private Sub CheckTheTwoTextBox(ByVal param_ArithmeticOperator As Char, ByVal param_ArithmeticSign As String)
        ' If the text of textbox called txtFirst is not empty and textbox called txtSecond is not visible 
        If Not txtFirst.Text = "" And txtSecond.Visible = False Then
            ' Invoke the User Defined Method called ClickedArithmeticButton and give all necessary values in arguments
            ClickedArithmeticButton(param_ArithmeticOperator, param_ArithmeticSign)
            ' If the text of textboxes called txtFirst and txtSecond are not both empty
        ElseIf Not txtFirst.Text = "" And Not txtSecond.Text = "" Then
            ' Invoke the User Defined Method called ClickedArithmeticButton()
            ClickedArithmeticButton(param_ArithmeticOperator, param_ArithmeticSign)
            ' Invoke the User Defined Method called CalculateTheAnswer()
            CalculateTheAnswer()
        End If
        ' Invoke the Object called 'sound' and use the Built-in Method called Play()
        sound.Play()

    End Sub
    Private Sub ClickedArithmeticButton(ByVal param_ArithmeticOperator As Char, ByVal param_ArithmeticSign As String)
        ' Set the arithmeticOperator from the value of parameter called param_ArithmeticOperator
        arithmeticOperator = param_ArithmeticOperator
        ' Get the text value of textbox called txtFirst and set in this field called firstGivenNumber
        firstGivenNumber = txtFirst.Text
        ' Add the current text value in textbox called txtFirst from the value of parameter called param_ArithmeticSign
        txtFirst.Text += param_ArithmeticSign
        ' Clear the text of this textbox called txtSecond
        txtSecond.Clear()
        ' Show this textbox called txtSecond
        txtSecond.Visible = True
    End Sub
    Private Sub Clear()
        txtFirst.Clear() ' Clear this textbox
        txtFirst.Visible = True ' Show this textbox
        txtSecond.Clear() ' Clear this textbox
        txtSecond.Visible = False ' Hide this textbox
        btnErase.Enabled = True ' Enable this button
        btnClearEntry.Enabled = True ' Enable this button
        btnPositiveOrNegative.Enabled = True ' Enable this button
        btnOne.Enabled = True ' Enable this button
        btnTwo.Enabled = True ' Enable this button
        btnThree.Enabled = True ' Enable this button
        btnFour.Enabled = True ' Enable this button
        btnFive.Enabled = True ' Enable this button
        btnSix.Enabled = True ' Enable this button
        btnSeven.Enabled = True ' Enable this button
        btnEight.Enabled = True ' Enable this button
        btnNine.Enabled = True ' Enable this button
        btnZero.Enabled = True ' Enable this button
        btnDot.Enabled = True ' Enable this button
        btnEqual.Enabled = True ' Enable this button
        ' Invoke the Object called 'sound' and use the Built-in Method called Play()
        sound.Play()

    End Sub
    Private Sub CalculateTheAnswer()
        Try
            ' If the text of textboxes called txtFirst and txtSecond are not both empty  
            If Not txtFirst.Text = "" And Not txtSecond.Text = "" Then
                ' Get the text value of textbox called txtSecond and set in this field called secondGivenNumber
                secondGivenNumber = txtSecond.Text
                ' Check the value of field called arithmeticOperator
                Select Case arithmeticOperator
                    ' If the value of field called arithmeticOperator is Plus
                    Case "+"
                        ' Then Add this two numbers
                        answer = firstGivenNumber + secondGivenNumber

                        ' If the value of field called arithmeticOperator is Minus
                    Case "-"
                        ' Then Subtract this two numbers
                        answer = firstGivenNumber - secondGivenNumber

                        ' If the value of field called arithmeticOperator is Divide
                    Case "/"
                        ' Then Divide this two numbers
                        answer = firstGivenNumber / secondGivenNumber

                        ' If the value of field called arithmeticOperator is Times
                    Case "*"
                        ' Then Multiply this two numbers
                        answer = firstGivenNumber * secondGivenNumber
                End Select

                ' If the value pf field called isbtnDotClicked is true
                If isbtnDotClicked = True Then
                    ' Then set the value of field called answer to the textbox called txtFirst
                    txtFirst.Text = answer

                    ' If the value of field called isbtnDotClicked is false
                Else
                    ' Convert the data type of field called answer, from double to integer
                    ' Then, Assign it to new local variable called convertToInt
                    Dim convertToInt As Integer = CInt(answer)
                    ' Then set the value of local variable called convertToInt to the textbox called txtFirst
                    txtFirst.Text = convertToInt
                End If

                ' Clear this textbox
                txtSecond.Clear()
                ' Hide this textbox
                txtSecond.Visible = False
                ' Set the boolean value of this field to false
                isbtnDotClicked = False

                ' If the text length of textbox called txtFirst is greater than or equal to 4
                If txtFirst.Text >= 4 Then
                    ' Declare a local variable for converting the format
                    Dim patternFormat As Double = Convert.ToDouble(txtFirst.Text)
                    ' If the textbox called txtFirst don't contains "."
                    If Not txtFirst.Text.Contains(".") Then
                        ' Convert the string format of the text of textbox called txtFirst 
                        txtFirst.Text = patternFormat.ToString("#,##0")
                    End If
                End If
            End If

            ' Invoke the Object called 'sound' and use the Built-in Method called Play()
            sound.Play()
        Catch ex As Exception ' I've declared here a General Exception
            ' Show the error in meaningful way and trace the error line
            MessageBox.Show(ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Sub
#End Region

End Class
