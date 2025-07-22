REM  *****  BASIC  *****
Option Explicit

private function GetCurRow as integer
	GetCurRow = ThisComponent.CurrentSelection.RangeAddress.StartRow
end function

private function GetSberClipboard() as string
	GetSberClipboard = ""

	dim cli as object, content as object, fla as object
	dim alen as integer, mime as object, tmime as object, ss as string
	cli = com.sun.star.datatransfer.clipboard.SystemClipboard.create()
	content = cli.getContents()
	fla = content.getTransferDataFlavors()
	' alen = fla.length
	
	ss = ""
	tmime = nothing
	for each mime in fla
		ss = mime.MimeType
		if ss = "text/plain;charset=utf-16" then
			tmime = mime
		end if
	next mime
	
	if not IsEmpty(tmime) then
		ss = content.getTransferData(tmime)
		GetSberClipboard = ss
	end if
	
end function

' Проверяет, является ли данный символ символом возврата каретки
private Function IsCRLF(c As String) As Boolean
    Dim CR As String: CR = Chr$(13)
    Dim LF As String: LF = Chr$(10)
    IsCRLF = (c = CR Or c = LF)
End Function

' Пропускает пробелы и символы возврата каретки - не далее, чем до конца строки
private Sub SkipCRLFSP(s As String, ByRef pos As Integer)
    Dim slen As Integer
    Dim sn As String
    slen = Len(s)
    Do While pos < slen
        sn = Mid(s, pos, 1) 's(pos)
        If Not (IsCRLF(sn) Or sn = " ") Then Exit Do
        pos = pos + 1
    Loop
End Sub

' Ищет вперед до конца строки или символов CR LF. Если не находит - возвращает конец строки
private Sub FindCRLF(s As String, ByRef pos As Integer)
    Dim slen As Integer
    slen = Len(s)
    Do While pos < slen
        If IsCRLF(Mid(s, pos, 1)) Then Exit Do
        pos = pos + 1
    Loop
End Sub

' Из строки s, полученной из буфера обмена, выделяет строки, игнорируя пустые строки
private sub CreateLinesArray(byref res() as string, s As String)
    Dim a() As String
    Dim i As Integer, n As Integer, st As Integer, en As Integer, slen As Long
    
    slen = Len(s)
    st = 1
    i = 0
    ReDim a(0 To 10000)
    
    Do While st < slen
        Call SkipCRLFSP(s, st)
        If st >= slen Then Exit Do
        en = st
        Call FindCRLF(s, en)
        If st <> en Then
            a(i) = Trim(Mid(s, st, en - st))
            i = i + 1
        End If
        st = en
    Loop

    n = i - 1
    ReDim Preserve a(0 To n)

	res = a
	
    'CreateLinesArray = a
    
End sub

' Ищет в массиве строку "Показать реквизиты" - где начинаются реквизиты
' Если не находит - возвращает -1
Function FindReqStart(a() As String) As Integer
    Dim k As Integer, n As Integer
    
    n = UBound(a) + 1
    
    For k = 0 To n - 1
        If a(k) = "Показать реквизиты" Then
            FindReqStart = k
            Exit Function
        End If
    Next k
    
    FindReqStart = -1
End Function


' Если от строки m вниз в четырех строках не пусто, то вставляет
' новую строку в позиции m
Sub WidenIfNotEmpty(sheet as object, m As Integer)
    Dim emp As Boolean
    Dim cell1, cell2 As object
    Dim i As Integer
    emp = true
    For i = 1 To 4
    	cell1 = sheet.getCellByPosition(0, m + i) 	' (col, row)
    	cell2 = sheet.getCellByPosition(1, m + i) 	' (col, row)
    	if cell1.Type <> com.sun.star.table.CellContentType.EMPTY or _
    	   cell2.Type <> com.sun.star.table.CellContentType.EMPTY then
            emp = False
            Exit For
        End If
    Next i
    
    If Not emp Then sheet.rows.insertByIndex(m, 1)
End Sub

' Тупо ставит перед строкой символ подчеркивания, а если строка заканчивается на
' " ?", то заменяет это на " Р"
Function Denumber(s As String) As String
    Dim n As Integer
    n = Len(s)
    If n >= 2 And Mid(s, n - 1, 2) = " ?" Then
    s = Mid(s, 1, n - 2) + " Р"
    End If
    Denumber = ". " + s

End Function


' Устанавливает свойства ячейки - жирность, выравнивание и пр
private sub SetCellProperties(cell as object)
	cell.CharWeight = 100
	cell.IsTextWrapped = true
    cell.HoriJustify = 0	' 0 - LEFT, 3 - RIGHT
    cell.VertJustify = 2	' 1 - UP, 2 - CENTER
end sub


' Отправляет результат в ячейки, начиная с элемента массива номер k
private Sub ToCells(a() As String, k As Integer)
    Dim n As Integer
    Dim i As Integer, st As Integer
    dim sheet as object, cell as object
    
    st = GetCurRow()
    n = UBound(a)
    
    sheet = ThisComponent.CurrentController.ActiveSheet	' текущий лист
    i = st
    Do While k <= n
        Call WidenIfNotEmpty(sheet, i)
        cell = sheet.getCellByPosition(0, i) 	' (col, row)
        SetCellProperties(cell)
        cell.string = a(k)
        k = k + 1
        If k > n Then Exit Do
        cell = sheet.getCellByPosition(1, i)	' (col, row)
        SetCellProperties(cell)
        cell.string = Denumber(a(k))
        i = i + 1
        k = k + 1
    Loop
   
End Sub

'---------------------------------------------------------------------
' Собственно вставка коммунального платежа из буфера обмена сбербанка
'---------------------------------------------------------------------
sub InsFromSber
	dim c as string
	c = GetSberClipboard()
	if Len(c) = 0 then
		MsgBox "Буфер обмена пуст", MB_ICONEXCLAMATION
		exit sub
	end if
	
	dim lines() as string
	CreateLinesArray(lines, c)
	
	Dim k As Integer
   
    k = FindReqStart(lines)
    If k < 0 Then
        MsgBox "Буфер обмена не содержит данные платежа от сбербанка", MB_ICONEXCLAMATION
    Else
        Call ToCells(lines, k + 1)
    End If
	
end sub


Sub Main
	dim cli as object, content as object, fla as object
	dim alen as integer, mime as object, tmime as object, ss as string
	cli = com.sun.star.datatransfer.clipboard.SystemClipboard.create()
	content = cli.getContents()
	fla = content.getTransferDataFlavors()
	' alen = fla.length
	
	ss = ""
	tmime = nothing
	for each mime in fla
		ss = mime.MimeType
		if ss = "text/plain;charset=utf-16" then
			tmime = mime
		end if
	next mime
	
	if not IsEmpty(tmime) then
		ss = content.getTransferData(tmime)
	end if
	
	dim a as integer
	a = 0

End Sub


