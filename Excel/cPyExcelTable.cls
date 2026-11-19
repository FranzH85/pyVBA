VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPyExcelTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const VBA_RowDelete As String = "VBA_RowDelete"

Public ExcelTable As cPyDicList
Public MyWorksheet As Worksheet
Public CurrentTableRange As Range

Sub TableRange(Range_ As Range, Optional FirstRowName_ As Boolean = True)
    Dim localExcelTableArray As Variant
    Dim firstColumnLetter  As String
    Dim firstColumnNumber As Long
    Dim columnIndex As Long
    Dim maxRowCount As Long
    Dim rowDelta As Long
    Dim keyName As String
    Dim I As Long
    Dim I2 As Long

    If Range_.Cells.Count = 1 Then
        ReDim localExcelTableArray(1 To 1, 1 To 1)
        localExcelTableArray(1, 1) = Range_.Value
    Else
        localExcelTableArray = Range_.Value
    End If

    Set ExcelTable = New cPyDicList
    Set MyWorksheet = Range_.Worksheet
    Set CurrentTableRange = Range_
    
    firstColumnNumber = Range_.Column
    columnIndex = UBound(localExcelTableArray, 2) - 1
    maxRowCount = UBound(localExcelTableArray, 1)
    
    rowDelta = Range_.Row - 1
    
    ExcelTable.SetNewMaxIndex columnIndex
    
    For I = 0 To columnIndex
        firstColumnLetter = Split(MyWorksheet.Cells(1, firstColumnNumber + I).Address, "$")(1)
        
        If FirstRowName_ Then
            keyName = localExcelTableArray(1, I + 1)
            ExcelTable.UpdateWithIndex I, keyName, New cPyDicList
            
            If maxRowCount >= 2 Then
                ExcelTable.ValueIndex(I).SetNewMaxIndex maxRowCount - 2
                For I2 = 2 To maxRowCount
                    ExcelTable.ValueIndex(I).UpdateWithIndex I2 - 2, rowDelta + I2, localExcelTableArray(I2, I + 1)
                Next I2
            Else
                ExcelTable.ValueIndex(I).SetNewMaxIndex -1
            End If
        Else
            ExcelTable.UpdateWithIndex I, firstColumnLetter, New cPyDicList
            ExcelTable.ValueIndex(I).SetNewMaxIndex maxRowCount - 1
            
            For I2 = 1 To maxRowCount
                ExcelTable.ValueIndex(I).UpdateWithIndex I2 - 1, rowDelta + I2, localExcelTableArray(I2, I + 1)
            Next I2
        End If
        ExcelTable.ValueIndex(I).MyTyp = firstColumnLetter
    Next I
End Sub

Sub TableRangeToEnd(Range_ As Range, Optional FirstRowName_ As Boolean = True)
    Dim lastColumnNumber As Long
    Dim lastRowNumber As Long
    Dim totalTableRange As Range

    lastColumnNumber = Range_.Worksheet.Cells(Range_.Row, Range_.Worksheet.Columns.Count).End(xlToLeft).Column
    lastRowNumber = Range_.Worksheet.Cells(Range_.Worksheet.Rows.Count, Range_.Column).End(xlUp).Row
    
    If lastColumnNumber < Range_.Column Then
        lastColumnNumber = Range_.Column
    End If
    
    If lastRowNumber < Range_.Row Then
        lastRowNumber = Range_.Row
    End If
    
    Set totalTableRange = Range_.Worksheet.Range(Range_, Range_.Worksheet.Cells(lastRowNumber, lastColumnNumber))
    
    TableRange totalTableRange, FirstRowName_
End Sub

Sub TableToRange(StartRange As Range, Optional FirstRowName_ As Boolean = True)
    Dim outputArray As Variant
    Dim maxColumnIndex As Long
    Dim maxRowIndex As Long
    Dim maxRowArrayIndex As Long
    Dim RowArrayI As Long
    Dim totalRows As Long
    Dim colIdx As Long
    Dim rowIdx As Long
    Dim arrayRowStart As Long
    Dim firstColumnNumber As Long
    Dim currentRowNumber As Long
    
    Set MyWorksheet = StartRange.Worksheet
    
    firstColumnNumber = StartRange.Column

    maxColumnIndex = ExcelTable.Imax
    maxRowIndex = ExcelTable.ValueIndex(0).Imax

    If FirstRowName_ Then
        totalRows = maxRowIndex + 2
        arrayRowStart = 2
    Else
        totalRows = maxRowIndex + 1
        arrayRowStart = 1
    End If
    
    maxRowArrayIndex = totalRows
    
    For rowIdx = 0 To maxRowIndex
        If ExcelTable.ValueIndex(0).Key(rowIdx) = VBA_RowDelete Then
            maxRowArrayIndex = maxRowArrayIndex - 1
        End If
    Next rowIdx

    ReDim outputArray(1 To maxRowArrayIndex, 1 To maxColumnIndex + 1)

    For colIdx = 0 To maxColumnIndex
        ExcelTable.ValueIndex(colIdx).MyTyp = Split(MyWorksheet.Cells(1, firstColumnNumber + colIdx).Address, "$")(1)
    
        If FirstRowName_ Then
            outputArray(1, colIdx + 1) = ExcelTable.Key(colIdx)
            currentRowNumber = StartRange.Row + 1
        Else
            currentRowNumber = StartRange.Row
        End If
        
        RowArrayI = arrayRowStart
        For rowIdx = 0 To maxRowIndex
            If ExcelTable.ValueIndex(0).Key(rowIdx) <> VBA_RowDelete Then
                outputArray(RowArrayI, colIdx + 1) = ExcelTable.ValueIndex(colIdx).ValueIndex(rowIdx)
                RowArrayI = RowArrayI + 1
                ExcelTable.ValueIndex(colIdx).Key(rowIdx) = currentRowNumber
                currentRowNumber = currentRowNumber + 1
            End If
        Next rowIdx
    Next colIdx

    StartRange.Worksheet.Range(StartRange, StartRange.Worksheet.Cells(StartRange.Row + maxRowArrayIndex - 1, StartRange.Column + maxColumnIndex)).Value = outputArray
    Set CurrentTableRange = StartRange.Worksheet.Range(StartRange, StartRange.Worksheet.Cells(StartRange.Row + maxRowArrayIndex - 1, StartRange.Column + maxColumnIndex))
End Sub

Sub CreateTable(ColumnsMaxIndex_ As Long, RowMaxIndex_ As Long)
    Dim I As Long
    Dim CurrentAddress As String
    
    Set ExcelTable = New cPyDicList
    ExcelTable.SetNewMaxIndex ColumnsMaxIndex_
    
    For I = 0 To ColumnsMaxIndex_
        CurrentAddress = Split(Cells(1, I + 1).Address, "$")(1)
        ExcelTable.UpdateWithIndex I, CurrentAddress, New cPyDicList
        ExcelTable.ValueIndex(I).SetNewMaxIndex RowMaxIndex_
    Next I
End Sub

Property Get ColumnName(ColumnIndex_ As Long) As String
    ColumnName = ExcelTable.Key(ColumnIndex_)
End Property
 
Property Let ColumnName(ColumnIndex_ As Long, Name_ As String)
    ExcelTable.Key(ColumnIndex_) = Name_
End Property

Property Get Value(ColumnName_ As String, RowIndex_ As Long) As Variant
Attribute Value.VB_UserMemId = 0
    Value = ExcelTable(ColumnName_).ValueIndex(RowIndex_)
End Property
 
Property Let Value(ColumnName_ As String, RowIndex_ As Long, Value_ As Variant)
Attribute Value.VB_UserMemId = 0
    ExcelTable(ColumnName_).ValueIndex(RowIndex_) = Value_
End Property

Property Get ImaxColumn() As Long
    ImaxColumn = ExcelTable.Imax
End Property

Property Get ImaxRow() As Long
    ImaxRow = ExcelTable.ValueIndex(0).Imax
End Property

Sub DeleteRow(Index_ As Long)
    ExcelTable.ValueIndex(0).Key(Index_) = VBA_RowDelete
End Sub

Function CellRange(ColumnName_ As String, RowIndex_ As Long) As Range
    Dim RangeString As String
    
    On Error GoTo ErrorHandler
    
    RangeString = ExcelTable(ColumnName_).MyTyp & ExcelTable(ColumnName_).Key(RowIndex_)
    Set CellRange = MyWorksheet.Range(RangeString)
    
    Exit Function
ErrorHandler:
     Err.Raise Number:=vbObjectError + 513, _
                  Source:="pyExcelTable.CellRange", _
                  Description:="Cell nicht gefunden!"
End Function

Function NameCellRange(ColumnName_ As String) As Range
    Dim RangeString As String
    
    On Error GoTo ErrorHandler
    
    RangeString = ExcelTable(ColumnName_).MyTyp & 1
    Set NameCellRange = MyWorksheet.Range(RangeString)
    
    Exit Function
ErrorHandler:
     Err.Raise Number:=vbObjectError + 513, _
                  Source:="pyExcelTable.CellRange", _
                  Description:="Namen Cell nicht gefunden!"
End Function

Sub ClearSheetRange(StartRange As Range, Optional FirstRowName_ As Boolean = True)
    Dim lastColumnNumber As Long
    Dim lastRowNumber As Long
    Dim totalTableRange As Range

    On Error GoTo ErrorHandler

    lastColumnNumber = StartRange.Worksheet.Cells(StartRange.Row, StartRange.Worksheet.Columns.Count).End(xlToLeft).Column
    lastRowNumber = StartRange.Worksheet.Cells(StartRange.Worksheet.Rows.Count, StartRange.Column).End(xlUp).Row

    If lastColumnNumber < StartRange.Column Then
        lastColumnNumber = StartRange.Column
    End If
    
    If lastRowNumber < StartRange.Row Then
        lastRowNumber = StartRange.Row
    End If

    If FirstRowName_ Then
        If lastRowNumber > StartRange.Row Then
            Set totalTableRange = StartRange.Worksheet.Range(StartRange.Offset(1, 0), StartRange.Worksheet.Cells(lastRowNumber, lastColumnNumber))
            totalTableRange.ClearContents
        End If
    Else
        Set totalTableRange = StartRange.Worksheet.Range(StartRange, StartRange.Worksheet.Cells(lastRowNumber, lastColumnNumber))
        totalTableRange.ClearContents
    End If

    Exit Sub
ErrorHandler:
    Err.Raise Number:=vbObjectError + 515, _
              Source:="pyExcelTable.ClearSheetRange", _
              Description:="Bereich auf dem Excel-Blatt konnte nicht geloescht werden!"
End Sub

Sub AddColumn(ColumnName_ As String, Optional AfterColumn As String = "")
    Dim I As Long
    Dim I2 As Long
    
    If ExcelTable.Key(ExcelTable.Imax) = AfterColumn Or AfterColumn = "" Then
        ExcelTable.Append ColumnName_, New cPyDicList
        ExcelTable.LastValue.SetNewMaxIndex ExcelTable.ValueIndex(0).Imax
    Else
        If ExcelTable.Exists(AfterColumn) Then
            ExcelTable.Append ExcelTable.Key(ExcelTable.Imax), ExcelTable.ValueIndex(ExcelTable.Imax)
            
            For I = ExcelTable.Imax - 1 To 1 Step -1
                If ExcelTable.Key(I - 1) = AfterColumn Then
                    ExcelTable.UpdateWithIndex I, ColumnName_, New cPyDicList
                    ExcelTable.ValueIndex(I).SetNewMaxIndex ExcelTable.ValueIndex(0).Imax
                    Exit For
                Else
                    ExcelTable.UpdateWithIndex I, ExcelTable.Key(I - 1), ExcelTable.ValueIndex(I - 1)
                End If
            Next I
        Else
            Err.Raise Number:=vbObjectError + 515, _
                  Source:="pyExcelTable.AddColumn", _
                  Description:="Spalte " & AfterColumn & " nicht vorhanden!"
        End If
    End If
    
    For I = 0 To ExcelTable.Imax
        ExcelTable.ValueIndex(I).MyTyp = Split(Cells(1, I + 1).Address, "$")(1)
        If ExcelTable.Key(I) = ColumnName_ Then
            For I2 = 0 To ExcelTable.ValueIndex(0).Imax
                ExcelTable.ValueIndex(I).Key(I2) = ExcelTable.ValueIndex(0).Key(I2)
            Next I2
        End If
    Next I
End Sub

Sub DeleteColumn(ColumnName_ As String)
    Dim I As Long
    
    If ExcelTable.Exists(ColumnName_) Then
        For I = ExcelTable.GetKeyIndex(ColumnName_) To ExcelTable.Imax - 1
            ExcelTable.UpdateWithIndex I, ExcelTable.Key(I + 1), ExcelTable.ValueIndex(I + 1)
        Next I
        
        ExcelTable.SetNewMaxIndex ExcelTable.Imax - 1
        
        For I = 0 To ExcelTable.Imax
            ExcelTable.ValueIndex(I).MyTyp = Split(Cells(1, I + 1).Address, "$")(1)
        Next I
    Else
        Err.Raise Number:=vbObjectError + 516, _
              Source:="pyExcelTable.DeleteColumn", _
              Description:="Spalte " & ColumnName_ & " nicht vorhanden!"
    End If
End Sub






