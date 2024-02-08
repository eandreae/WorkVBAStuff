Sub update_GamesListDates()
    ' The purpose of this Subroutine is meant to update the 2 by X list of games.

    ' Variable Declarations
    Dim InputRow As Integer
    Dim InputCol As Integer
    Dim InputColRange As String
    Dim InputGameRowIncrement As Integer

    Dim TargetRow As Integer
    Dim TargetCol As Integer
    Dim TargetColRange As String
    Dim TargetGameRowIncrement As Integer

    Dim OnLeft As Boolean

    Dim NextInputGameOffset As Integer
    Dim TargetBackToTop As Integer
    Dim LeftRightOffset As Integer
    Dim NextOutputGame As Integer

    ' Initial values for variables
    InputRow = 5              ' Row 5 is where the data begins
    InputCol = 19             ' Col 19 is where the data begins
    InputColRange = "S"       ' R is Col 19, where the data begins
    InputGameRowIncrement = 5 ' Starts at 5, since 5 is where the data begins

    TargetRow = 6              ' Row 6 is where the output data begins
    TargetCol = 9              ' Col 9 is where the output data begins
    TargetColRange = "I"       ' I is Col 9, where the output data begins
    TargetGameRowIncrement = 5 ' Starts at 5, since 5 is where the output data begins

    OnLeft = True ' Initially set to true, since outputting to the game on the left.

    NextInputGameOffset = 18
    TargetBackToTop = 5
    LeftRightOffset = 4
    NextOutputGame = 5


    ' Step 1, read the relevant information.
    ' List of Relevant Information:
        ' 1. Pre-Production Review - Location = Header + 4
        Dim PrePro_Offset As Integer
        PrePro_Offset = 4
        ' 2. Handoff to SQA        - Location = Pre-Pro + 3 (Header + 7)
        Dim Ho2SQA_Offset As Integer
        Ho2SQA_Offset = 7
        ' 3. 1st Paytables         - Location = H02SQA + 3 (Header + 10)
        Dim Paytables_Offset As Integer
        Paytables_Offset = 10
        ' 4. Polish Complete       - Location = 1st Paytables + 2 (Header + 12)
        Dim Polish_Offset As Integer
        Polish_Offset = 12
        ' 5. Golden Game           - Location = Polish Complete + 3 (Header + 15)
        Dim Golden_Offset As Integer
        Golden_Offset = 15
        ' 6. Handoff to Compliance - Location = Golden Game + 1 (Header + 16)
        Dim Ho2C_Offset As Integer
        Ho2C_Offset = 16

    ' Increment through each game on the Full List, outputting onto the Game List
    ' In-between game increment = 18

    ' While there is a game on the Full List to Read data from
    While Len(Range(InputColRange & CStr(InputGameRowIncrement)).Value) > 0
        ' Starts with the header at row 5. Assume initial state is in an accurate place.
        InputRow = InputGameRowIncrement

        ' Read the Pre-Production Review Gate, write it to the target location.
        Cells(TargetRow, TargetCol).Value = Cells(InputRow + PrePro_Offset, InputCol).Value
        ' Read whether or not it's closed, write it to the target location.
        Cells(TargetRow, TargetCol+1).Value = Cells(InputRow + PrePro_Offset, InputCol+1).Value
        ' Increment the Target by 1 Row.
        TargetRow = TargetRow + 1

        ' Read the Handoff to SQA gate, write it to the target location.
        Cells(TargetRow, TargetCol).Value = Cells(InputRow + Ho2SQA_Offset, InputCol).Value
        ' Read whether or not it's closed, write it to the target location.
        Cells(TargetRow, TargetCol+1).Value = Cells(InputRow + Ho2SQA_Offset, InputCol+1).Value
        ' Increment the Target by 1 Row.
        TargetRow = TargetRow + 1

        ' Read the 1st Paytable gate, write it to the target location.
        Cells(TargetRow, TargetCol).Value = Cells(InputRow + Paytables_Offset, InputCol).Value
        ' Read whether or not it's closed, write it to the target location.
        Cells(TargetRow, TargetCol+1).Value = Cells(InputRow + Paytables_Offset, InputCol+1).Value
        ' Increment the Target by 1 Row.
        TargetRow = TargetRow + 1

        ' Read the Polish Complete gate, write it to the target location.
        Cells(TargetRow, TargetCol).Value = Cells(InputRow + Polish_Offset, InputCol).Value
        ' Read whether or not it's closed, write it to the target location.
        Cells(TargetRow, TargetCol+1).Value = Cells(InputRow + Polish_Offset, InputCol+1).Value
        ' Increment the Target by 1 Row.
        TargetRow = TargetRow + 1

        ' Read the Golden Game gate, write it to the target location.
        Cells(TargetRow, TargetCol).Value = Cells(InputRow + Golden_Offset, InputCol).Value
        ' Read whether or not it's closed, write it to the target location.
        Cells(TargetRow, TargetCol+1).Value = Cells(InputRow + Golden_Offset, InputCol+1).Value
        ' Increment the Target by 1 Row.
        TargetRow = TargetRow + 1

        ' Read the Handoff to Compliance gate, write it to the target location.
        Cells(TargetRow, TargetCol).Value = Cells(InputRow + Ho2C_Offset, InputCol).Value
        ' Read whether or not it's closed, write it to the target location.
        Cells(TargetRow, TargetCol+1).Value = Cells(InputRow + Ho2C_Offset, InputCol+1).Value
        
        ' Increment the InputGameRowIncrement by 18 to move it to the next game.
        InputGameRowIncrement = InputGameRowIncrement + NextInputGameOffset

        ' Increment the Target position to the next game position.
        ' The next game position depends if the previous target was on the left or right.
        If OnLeft Then
            ' If the target was just on the left, increment it to the right.

            ' Move the target Column to the right by the LeftRightOffset
            TargetCol = TargetCol + LeftRightOffset

            ' Move the target Row up to the top
            TargetRow = TargetRow - TargetBackToTop

            ' Flip OnLeft to False, as now we're on the right.
            OnLeft = False

        Else 
            ' If the target was NOT on the left, incremement it to the left.

            ' Move the target Row down by the NextOutputGame offset
            TargetRow = TargetRow + NextOutputGame

            ' Move the target Column to the left by the LeftRightOffset
            TargetCol = TargetCol - LeftRightOffset

            ' Flip OnLeft to True, as we're on the left.
            OnLeft = true

        End If

    Wend

End Sub