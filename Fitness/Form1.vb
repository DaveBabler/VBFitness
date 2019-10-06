' Program Name: FitnessChallenge
' Author:       Dave Babler
' Date:         2019-10-03
' Purpose:      Program collects then displays a team's individual weight loss numbers
'               then the program calculates and displays the average (MEAN) weight loss for the entire team.
Option Strict On
Public Class frmFitness



    Private Sub frmFitness_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub btnWeightLoss_Click(sender As Object, e As EventArgs) Handles btnWeightLoss.Click
        'Button calculates up to 8 weight loss values and then gets the mathematical mean for the team.

        'declare math & user value  holding variables
        Dim strWeightLoss As String
        Dim decWeightLoss As Decimal
        Dim decAverageLoss As Decimal
        Dim decTotalWeightLoss As Decimal

        'Variables for the InputBox Function
        Dim strInputMessage As String = "Enter the weight loss for the team member"
        Dim strInputHeading As String = "Weight Loss"
        Dim strNormalMessage As String = "Enter the weight loss for team member #"
        Dim strNonNumericError As String = "Error - Enter a number for the weight loss of team member #"
        Dim strNegativeError As String = "Error - Enter a positive number for the weight loss of team member #"

        'Declare Loop Control Variables
        Dim strCancelClicked As String = ""
        Dim intMaxNumberOfEnries As Integer = 8
        Dim intNumberOfEntries As Integer = 1

        'loop allows users to enter weight loss of up to 8 team membmers
        ' terminates when the user has entered 8 values or clicks the cancel or cluse  button in the Input box
        strWeightLoss = InputBox(strInputMessage & intNumberOfEntries, strInputHeading, " ")
        'Note on the space at the end of the Input Box, if the person enters nothing in the box then clicks okit enters a space
        'so it is treated differently as clicking the cancel button

        Do Until intNumberOfEntries > intMaxNumberOfEnries Or strWeightLoss = strCancelClicked
            If IsNumeric(strWeightLoss) Then
                decWeightLoss = Convert.ToDecimal(strWeightLoss)
                If decWeightLoss > 0 Then
                    lstWeightLoss.Items.Add(decWeightLoss)
                    decTotalWeightLoss += decWeightLoss
                    intNumberOfEntries += 1
                    strInputMessage = strNormalMessage

                Else
                    strInputMessage = strNegativeError
                End If

            Else
                strInputMessage = strNonNumericError
            End If

            If intNumberOfEntries <= intMaxNumberOfEnries Then
                strWeightLoss = InputBox(strInputMessage & intNumberOfEntries, strInputHeading, " ")
            End If
        Loop
        'Calculate and display average team weightloss
        If intNumberOfEntries > 1 Then
            lblAverageLoss.Visible = True
            decAverageLoss = decTotalWeightLoss / (intNumberOfEntries - 1)  ' the loop adds one so we must subtract one
            lblAverageLoss.Text = "Average weight loss for the team is " &
            decAverageLoss.ToString("F1") & " lbs"

        Else
            MsgBox("No weight loss value entered")

        End If
        ' Disables the Weight Loss button
        btnWeightLoss.Enabled = False

    End Sub
End Class
