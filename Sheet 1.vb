Public Sub ngijoCoy()
     If TypeName(Selection) = "Range" Then
        Dim rng As Range, cell As Range
        Dim cari As String
        Dim asd As String
        
        Set rng = Selection
        cari = "dimension"
        For Each cell In rng
            
            If (IsEmpty(Cells(cell.Cells.Row, cell.Cells.Column).Value) = False) Then
                asd = Cells(cell.Cells.Row, cell.Cells.Column).Text
                
                If InStr(1, asd, cari, vbTextCompare) > 0 Then
                    Cells(cell.Cells.Row, cell.Cells.Column).Interior.ColorIndex = 4
                    Cells(cell.Cells.Row, cell.Cells.Column + 1).Interior.ColorIndex = 4
                End If
            End If
        Next cell
     End If
End Sub

Public Sub ngapusBukanIjoCoy()
     If TypeName(Selection) = "Range" Then
        Dim rng As Range, cell As Range
        Dim cari As String
        Dim asd As String
        
        Set rng = Selection
        cari = "dimension"
        For Each cell In rng
            
            If (IsEmpty(Cells(cell.Cells.Row, cell.Cells.Column).Value) = False) Then
                If Cells(cell.Cells.Row, cell.Cells.Column).Interior.ColorIndex <> 4 Then
                    Cells(cell.Cells.Row, cell.Cells.Column).ClearContents
                End If
            End If
        Next cell
     End If
End Sub


Public Sub ngapusIjoCoy()
     If TypeName(Selection) = "Range" Then
        Dim rng As Range, cell As Range
        Dim cari As String
        Dim asd As String
        
        Set rng = Selection
        cari = "dimension"
        For Each cell In rng
            
            If (IsEmpty(Cells(cell.Cells.Row, cell.Cells.Column).Value) = False) Then
                If Cells(cell.Cells.Row, cell.Cells.Column).Interior.ColorIndex = 4 Then
                    Cells(cell.Cells.Row, cell.Cells.Column).ClearContents
                    Cells(cell.Cells.Row, cell.Cells.Column).Interior.ColorIndex = 0
                End If
            End If
        Next cell
     End If
End Sub
