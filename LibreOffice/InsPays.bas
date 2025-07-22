REM  *****  BASIC  *****
option explicit

' Вычисляет текущую строку, на которой стоит курсор
private function GetCurRow as integer
	GetCurRow = ThisComponent.CurrentSelection.RangeAddress.StartRow
end function

' Ищет глубину квитанции - от начала до строки --------
private function FindDepth as integer
	dim r as integer, n as integer
	dim s as string
	dim sheet as object, cell as object

	sheet = ThisComponent.CurrentController.ActiveSheet	' текущий лист
	n = 200
	for r = 1 to n
		cell = sheet.getCellByPosition(0, r)	' (col, row)
		if cell.Type = com.sun.star.table.CellContentType.TEXT then
			s = cell.getString()
			if Len(s) > 10 and Mid(s, 2, 10) = "----------" then
				FindDepth = r + 1
				exit function
			end if
		end if
	next r

	FindDepth = -1
end function


' Из строки вида 2024 01 20 Труда 49 оставляет только Труда 49
private function CutDate(s as string) as string
    dim c as string, i as integer
    for i = 1 to Len(s)
       c = Mid(s, i, 1)
       if c <> " " and (c < "0" Or c > "9") then exit for
    next i
    if i < Len(s) then c = Mid(s, i) else c = ""
    CutDate = c
end function


' откусывает находящуюся в начале дату и ставит сегодняшнюю
private function UpdateHeader(h as string) as string
    Dim y, m, d, da, s
    da = Now
    y = Year(da)
    m = Month(da)
    d = Day(da)
    s = Format(y, "####") + " " + Format(m, "0#") + " " + Format(d, "0#") + " " + CutDate(h)
    UpdateHeader = s
end function

' вставляет минусы в текущей строке. И выбирает ячейку А в следующей строке
' возвращает номер следующей после минусов строки
private function minuses as integer
	dim r as integer
	r = GetCurRow
	dim sheet as object, cell as object
	sheet = ThisComponent.CurrentController.ActiveSheet	' текущий лист
	cell = sheet.getCellByPosition(0, r)	' (col, row)
	cell.string = "'--------------------------------------------------------------------------------------------"
	' cell = sheet.getCellByPosition(0, r + 1)
	' ThisComponent.CurrentController.Select(cell)
	minuses = r + 1
end function


'=========================================================================

' Вставляет пустые строки количеством от начала до --------, умноженное на 1.5
sub Ins
	' MsgBox "Ins - Временно не реализовано"
	Dim res as integer, sheet as object, cell as object, ss as string
    res = FindDepth()
    If res <= 0 Then
        MsgBox ("Не найдена строка ------------- в конце")
        Exit Sub
    End If
    sheet = ThisComponent.CurrentController.ActiveSheet	' текущий лист
    cell = sheet.getCellByPosition(0, 0)	' (col, row)
	if cell.Type <> com.sun.star.table.CellContentType.TEXT then
		MsgBox "В первой ячейке первой строки ожидался текст, начинающийся с даты"
		exit sub
	end if
    ss = cell.string
    ss = UpdateHeader(ss)
    res = res * 3
    res = res / 2
    sheet.rows.insertByIndex(0, res)	' вставка в начало таблицы пустых строк
    cell = sheet.getCellByPosition(0, 0)	' (col, row)
    cell.string = ss	' UpdateHeader(ss)
    cell.CharWeight = 150	' жирный шрифт
end sub

' Вставляет пустые строки количеством от начала 5 штук - с копированием заголовка
sub Ins5
    Dim res as integer, cell as object, sheet as object, s as string
    sheet = ThisComponent.CurrentController.ActiveSheet	' текущий лист
    cell = sheet.getCellByPosition(0, 0)	' (col, row)
	if cell.Type <> com.sun.star.table.CellContentType.TEXT then
		MsgBox "В первой ячейке первой строки ожидался текст, начинающийся с даты"
		exit sub
	end if
	s = UpdateHeader(cell.string)
    sheet.rows.insertByIndex(0, 5)	' вставка в начало таблицы пустых строк
    cell = sheet.getCellByPosition(0, 0)	' (col, row)
    cell.string = s		' UpdateHeader(s)
    cell.CharWeight = 150	' жирный шрифт
end sub

' Вставляет одну пустую строку в текущем положении курсора
sub Ins1
	dim sheet as object
	sheet = ThisComponent.CurrentController.ActiveSheet	' текущий лист
	dim arow as integer
	arow = GetCurRow()
	sheet.rows.insertByIndex(arow, 1)
end sub

sub MSqueeze
	dim start as integer, r as integer, c as integer, h as integer
	dim sheet as object, cell as object
	start = minuses()
	
	' ищем непустую строку - непустая, если что-то есть в первых 10 колонках
	sheet = ThisComponent.CurrentController.ActiveSheet	' текущий лист
	for r = start to start + 500
		for c = 0 to 10
			cell = sheet.getCellByPosition(c, r)	' (col, row)
			if cell.type <> com.sun.star.table.CellContentType.EMPTY then: exit for: end if
		next c
		if cell.type <> com.sun.star.table.CellContentType.EMPTY then	' нашли непустую строку
			h = r - start - 2
			if h > 0 then
				sheet.rows.RemoveByIndex(start, h)
			end if
			exit sub
		end if
	next r
	
	MsgBox "Отсюда вниз не найдена непустая строка"
	
end sub


' Удаление одной текущей строки
sub Del1
	' MsgBox "Del1 - Временно не реализовано"
	dim sheet as object
	sheet = ThisComponent.CurrentController.ActiveSheet
	dim arow as integer
	arow = GetCurRow
	sheet.rows.RemoveByIndex(arow, 1)

end sub
