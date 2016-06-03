Public Module ExcelUtilityFunctions

    'for editing. saves the coordinates of the point being edited so if the edit operation is cancelled, the original coordinates can be restored
    Friend tempPosX, tempPosY, tempPosZ
    'this is a flag to indicate the teaching step number of the current teach process i.e. line needs to be taught 2 points, arc 3 points.
    Friend m_TeachStepNumber As Integer = 1
    'this is a flag for whether the user is in the middle of editing
    Friend m_EditStateFlag As Boolean = False
    'this is a flag for whether the user is in the middle of teaching
    Friend m_ProgrammingStateFlag As Boolean = False
    'this is a flag for deciding whether to update the machine coordinates in the spreadsheet while teaching
    Friend m_PosUpdate As Boolean = False
    'this contains the command currently being taught/edited
    Friend type As String = ""

    Friend m_SteppingPostFlag As Boolean = False
    'flags
    Friend m_Row As Integer = 2
    Friend m_InLinkRange As Boolean = False
    Friend m_ArrayRow As Integer = 0
    Friend m_Column As Integer = 2

    Public Sub EnableCoordinateUpdateInSpreadsheet()
        m_PosUpdate = True
    End Sub

    Public Sub DisableCoordinateUpdateInSpreadsheet()
        m_PosUpdate = False
    End Sub

#Region "Excel GUI"

    Public Sub EnableReferenceCommandBlock()
        Dim i As Integer
        For i = 0 To gMaxReferButtons
            Programming.ReferenceCommandBlock.Buttons(i).Enabled = True
        Next
        TraceGCCollect
    End Sub

    Public Sub DisableReferenceCommandBlock()
        Dim i As Integer
        For i = 0 To gMaxReferButtons
            Programming.ReferenceCommandBlock.Buttons(i).Enabled = False
        Next
        TraceGCCollect
    End Sub

    Public Sub EnableReferenceCommandBlockButton(ByVal CmdButton As Integer)
        Programming.ReferenceCommandBlock.Buttons(CmdButton - 1).Enabled = True
        TraceGCCollect
    End Sub

    Public Sub EnableElementsCommandBlockButton(ByVal CmdButton As Integer)
        Programming.ElementsCommandBlock.Buttons(CmdButton - 1).Enabled = True
        TraceGCCollect
    End Sub

    Public Sub DisableElementsCommandBlockButton(ByVal CmdButton As Integer)
        Programming.ElementsCommandBlock.Buttons(CmdButton - 1).Enabled = False
        TraceGCCollect
    End Sub

    Public Sub EnableElementsCommandBlock()
        Dim i As Integer
        For i = 0 To gMaxElementButtons
            Programming.ElementsCommandBlock.Buttons(i).Enabled = True
            If i = 9 Then Programming.ElementsCommandBlock.Buttons(i).Enabled = False 'disable move command for now
        Next
        TraceGCCollect
    End Sub

    Public Sub DisableElementsCommandBlock()
        Dim i As Integer
        For i = 0 To gMaxElementButtons
            Programming.ElementsCommandBlock.Buttons(i).Enabled = False
        Next
        TraceGCCollect
    End Sub

    Public Sub DisableTeachingButtons()
        DisableReferenceCommandBlock()
        DisableElementsCommandBlock()
    End Sub

    Public Sub EnableTeachingButtons()
        EnableReferenceCommandBlock()
        EnableElementsCommandBlock()
    End Sub

    Public Sub DisableTeachingToolbar()
        Programming.TeachingToolbar.Buttons(0).Enabled = False
        Programming.TeachingToolbar.Buttons(1).Enabled = False
    End Sub

    Public Sub EnableTeachingToolbar()
        Programming.TeachingToolbar.Buttons(0).Enabled = True
        Programming.TeachingToolbar.Buttons(1).Enabled = True
    End Sub

    Public Sub EnableEditingToolbar()
        Programming.EditingToolbar.Buttons(0).Enabled = True
        Programming.EditingToolbar.Buttons(1).Enabled = True
    End Sub

    Public Sub DisableEditingToolbar()
        Programming.EditingToolbar.Buttons(0).Enabled = False
        Programming.EditingToolbar.Buttons(1).Enabled = False
    End Sub

    Public Sub EnableEditingToolbar(ByVal CmdButton As Integer, ByVal flag As Boolean)
        Programming.EditingToolbar.Buttons(CmdButton).Enabled = flag
    End Sub

    Public Sub EnableEditingToolbarSwitchButton()
        Programming.EditingToolbar.Buttons(0).Enabled = True
    End Sub

    Public Sub DisableEditingToolbarSwitchButton()
        Programming.EditingToolbar.Buttons(0).Enabled = False
    End Sub

    Public Sub EnableEditingToolbarEditButton()
        Programming.EditingToolbar.Buttons(1).Enabled = True
    End Sub

    Public Sub DisableEditingToolbarEditButton()
        Programming.EditingToolbar.Buttons(1).Enabled = False
    End Sub

    Public Sub EnableTeachingToolbarOKButton()
        Programming.TeachingToolbar.Buttons(0).Enabled = True
    End Sub

    Public Sub DisableTeachingToolbarOKButton()
        Programming.TeachingToolbar.Buttons(0).Enabled = False
    End Sub

    Public Sub EnableTeachingToolbarCancelButton()
        Programming.TeachingToolbar.Buttons(1).Enabled = True
    End Sub

    Public Sub DisableTeachingToolbarCancelButton()
        Programming.TeachingToolbar.Buttons(1).Enabled = False
    End Sub

    Public Sub DisplaySpreadsheetTabs()
        Programming.AxSpreadsheetProgramming.DisplayWorkbookTabs = True
    End Sub

    Public Sub HideSpreadsheetTabs()
        Programming.AxSpreadsheetProgramming.DisplayWorkbookTabs = False
    End Sub

    Sub ToggleButtonsForTeachingStart()
        EnableTeachingToolbar()
        DisableTeachingButtons()
        DisableEditingToolbar()
    End Sub

    Sub ToggleButtonsForTeachingStop()
        DisableTeachingToolbar()
        EnableTeachingButtons()
    End Sub

    Function GetCell(ByVal row As Integer, ByVal column As Integer)
        Try
            Return Programming.AxSpreadsheetProgramming.ActiveSheet.Cells(row, column)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Function GetCellValue(ByVal row As Integer, ByVal column As Integer)
        Try
            Return Programming.AxSpreadsheetProgramming.ActiveSheet.Cells(row, column).value
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Sub SetCellValue(ByVal row As Integer, ByVal column As Integer, ByVal value As Object)
        'Do
        '    MySleep(10)
        '    TraceDoEvents()
        'Loop Until Not m_PosUpdate
        Try
            Programming.AxSpreadsheetProgramming.ActiveSheet.Cells(row, column) = value
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Sub SetActiveCellValue(ByVal value As Object)
        Try
            Programming.AxSpreadsheetProgramming.ActiveCell.Value = value
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Function GetActiveSheetName()
        Return Programming.AxSpreadsheetProgramming.ActiveSheet.Name
    End Function

    Function GetActiveCellValue()
        Try
            Return Programming.AxSpreadsheetProgramming.ActiveCell.Value
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Function GetActiveCellRow()
        Try
            Return Programming.AxSpreadsheetProgramming.ActiveCell.Row
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Function GetActiveCellColumn()
        Try
            Return Programming.AxSpreadsheetProgramming.ActiveCell.Column
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Function GetSheetCount()
        Try
            Return Programming.AxSpreadsheetProgramming.ActiveWorkbook.Worksheets.Count()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Sub ClearRow(ByVal row As Integer)
        Try
            Programming.AxSpreadsheetProgramming.ActiveSheet.Rows(row).Clear()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Sub DeleteRow(ByVal row As Integer)
        Try
            Programming.AxSpreadsheetProgramming.ActiveSheet.Rows(row).Delete()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Sub UpdateSpreadsheet()
        Try
            Programming.AxSpreadsheetProgramming.Update()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

#End Region

End Module
