REM  *****  BASIC  *****
Option Explicit

' Если в пределах от 0 до 100 строк нашлась строка --------------------
' то возвращает высоту, иначе возвращает -1
private function CountMonthHeight(sn as integer) as integer
	dim r as integer, n as integer
	dim s as string
	dim sheet as object, cell as object

	sheet = ThisComponent.Sheets(sn)
	n = 100
	for r = 1 to n
	cell = sheet.getCellByPosition(0, r)	' (col, row)
		if cell.Type = com.sun.star.table.CellContentType.TEXT then
			s = cell.getString()
			if Len(s) > 10 and Mid(s, 1, 10) = "----------" then
				CountMonthHeight = r + 1
				exit function
			end if
		end if
	next r

	CountMonthHeight = -1
end function


' Вставляет пустые строки в начало листа, потом копирует в него 
' содержимое предыдущего месяца
' sn - номер листа, mh - высота месяца - на один больше номера последней строки
private sub InsCopyPrevMonth(sn as integer, mh as integer)
	dim sheet as object
	sheet = ThisComponent.Sheets(sn)
	sheet.rows.insertByIndex(0, mh)
	
	dim CellRangeAddress as New com.sun.star.table.CellRangeAddress
	dim CellAddress as new com.sun.star.table.CellAddress
	
	CellRangeAddress.Sheet = sn
	CellRangeAddress.StartColumn = 0
	CellRangeAddress.StartRow = mh
	CellRangeAddress.EndColumn = 30
	CellRangeAddress.EndRow = mh + mh - 1
	
	CellAddress.Sheet = sn
	CellAddress.Column = 0
	CellAddress.Row = 0
	
	sheet.copyRange(CellAddress, CellRangeAddress)
	
end sub

private function CreateCurMonthYear as string
	dim dat as Date
	dim mon as string
	dat = Now
	select case Month(dat)
		case 1: mon = "Январь"
		case 2: mon = "Февраль"
		case 4: mon = "Март"
		case 4: mon = "Апрель"
		case 5: mon = "Май"
		case 6: mon = "Июнь"
		case 7: mon = "Июль"
		case 8: mon = "Август"
		case 9: mon = "Сентябрь"
		case 10: mon = "Октябрь"
		case 11: mon = "Ноябрь"
		case 12: mon = "Декабрь"
		case else: mon = "Undefined"
	end select
	CreateCurMonthYear = mon + " " + Trim(Str(Year(dat)))
	
end function


' является ли символ snv цифрой. Если да, то nv1 - ее значение
private function IsDigit(byref nv1 as double, snv as string) as boolean
	IsDigit = false
	if snv = "0" then: nv1 = 0
	elseif snv = "1" then nv1 = 1
	elseif snv = "2" then nv1 = 2
	elseif snv = "3" then nv1 = 3
	elseif snv = "4" then nv1 = 4
	elseif snv = "5" then nv1 = 5
	elseif snv = "6" then nv1 = 6
	elseif snv = "7" then nv1 = 7
	elseif snv = "8" then nv1 = 8
	elseif snv = "9" then nv1 = 9
	else: exit function
	end if
	IsDigit = true
end function


' Проверяет, содержит ли ячейка указание на замену типа "E1234.56"
' Если да, то nc - номер колонки (Е - 4), nv - значение (1234.56)
private function IsReplaceCell(byref nc as integer, byref nv as double, cell as object) as boolean
	dim s as string
	IsReplaceCell = false
	
	' если в ячейки - не текст, или слишком короткий текст - возвращаем false
	if cell.type <> com.sun.star.table.CellContentType.TEXT then exit function
	s = cell.string
	if len(s) < 2 then exit function
	
	' если текст начинается не с правильной буквы - возвращаем false
	dim snc as string
	snc = Mid(s, 1, 1)
	if snc = "B" or snc = "b" then: nc = 1
	elseif snc = "C" or snc = "c" then: nc = 2
	elseif snc = "D" or snc = "d" then: nc = 3
	elseif snc = "E" or snc = "e" then: nc = 4
	elseif snc = "F" or snc = "f" then: nc = 5
	else: exit function
	end if
	
	' штош... пытаемся прочитать число после буквы, до десятичной точки
	dim pos as integer, alen as integer, snv as string, nv1 as double
	dim start as integer
	nv = 0
	alen = len(s)
	start = 2
	for pos = start to alen
		snv = Mid(s, pos, 1)
		if not IsDigit(nv1, snv) then: exit for: end if
		nv = nv * 10 + nv1
	next pos
	if pos = start then: exit function: end if	' если не было ни одной цифры - фигня
	if pos > alen then			' если конец строки - целое число
		IsReplaceCell = true
		exit function
	end if
	if snv <> "." and snv <> "," then: exit function: end if 	' если не дес точка - ошибка
	
	' читаем цифры после запятой
	dim dec as double: dec = 1
	start = pos + 1
	for pos = start to alen
		snv = Mid(s, pos, 1)
		if not IsDigit(nv1, snv) then: exit function: end if
		dec = dec / 10
		nv = nv + nv1 * dec
	next pos
	if pos = start then: exit function: end if	' если не было ни одной цифры - фигня
	if pos >= alen then			' если конец строки - целое число
		IsReplaceCell = true
		exit function
	end if
	
end function


' проверка единичного случая IsReplaceCellUnitTest
' при этом портит ячейку (0,0)
private function ircut(s as string, res as boolean, nc as integer, nv as double) as boolean
	dim ares as boolean, anc as integer, anv as double
	dim doc as object, sheet as object, cell as object
	doc = ThisComponent
	sheet = doc.sheets(0)
	cell = sheet.getCellByPosition(0, 0)
	cell.string = s
	ares = IsReplaceCell(anc, anv, cell)
	cell.string = ""
	if ares <> res then: exit function: end if
	if ares = false then: ircut = true: exit function: end if
	ircut = nc = anc and abs(nv - anv) < 0.00000001
end function

' Тестирование IsReplaceCell
private sub IsReplaceCellUnitTest
	if not ircut("E123,45", true, 4, 123.45) then: goto Bad: end if
	if not ircut("E", false, 0, 0) then: goto Bad: end if
	if not ircut("B.", false, 0, 0) then: goto Bad: end if
	if not ircut("b12", true, 1, 12) then: goto Bad: end if
	if not ircut("C12.", false, 1, 12) then: goto Bad: end if
	if not ircut("C0.12", true, 2, 0.12) then: goto Bad: end if
	if not ircut("C0.12.", false, 2, 0.12) then: goto Bad: end if
	if not ircut("c1.23", true, 2, 1.23) then: goto Bad: end if
	if not ircut("D2.34", true, 3, 2.34) then: goto Bad: end if
	if not ircut("d3.4", true, 3, 3.4) then: goto Bad: end if
	if not ircut("E4.56", true, 4, 4.56) then: goto Bad: end if
	if not ircut("e5.67", true, 4, 5.67) then: goto Bad: end if
	if not ircut("F6.78", true, 5, 6.78) then: goto Bad: end if
	if not ircut("f7.89", true, 5, 7.89) then: goto Bad: end if
	
	MsgBox "IsReplaceCellUnitTest успешно пройден"
	exit sub
	
Bad: 
	MsgBox "IsReplaceCellUnitTest обнаружил ошибку"
end sub


' вставляет в файл показаний новый месяц
sub InsNextMonth
	dim hdr as string, nam as string, ask as string

	'dim cc as ScCellObj
	
	hdr = CreateCurMonthYear
	
	dim sn as integer, n as integer, sh as integer, answ as integer
	dim Doc As Object, sheet as object, cell as object
	doc = ThisComponent
	n = doc.sheets.count
	
	for sn = 0 to n - 1
		sheet = doc.sheets(sn)
		nam = sheet.name
		cell = sheet.getCellByPosition(1, 0)	' (col, row)
		
		' Если вообще нет заголовка в виде месяца и года
		if cell.type <> com.sun.star.table.CellContentType.TEXT then
			ask = "На листе '" + nam + "' некорректный заголовок" + chr(10) + chr(12)
			ask = ask + "Ожидалось название месяца и год" + chr(10) + chr(12)
			ask = ask + "OK - перейти к следующему листу, Отмена - прекратить"
			answ = MsgBox(ask, MB_OKCANCEL)
			if answ = IDOK then goto nextsn
			exit for
		end if
		
		' если на листе уже есть заголовок с этим месяцем - годом
		ask = cell.string
		if ask = hdr then
			ask = "На листе '" + nam + "' уже есть заголовок '" 
			ask = ask + hdr + "'" + chr(10) + chr(12)
			ask = ask + "OK - перейти к следующему листу, Отмена - прекратить"
			answ = MsgBox(ask, MB_OKCANCEL)
			if answ = IDOK then goto nextsn
			exit for
		end if
		
		' если на листе нет окончания текущего месяца -------
		sh = CountMonthHeight(sn)
		if sh < 0 then
			ask = "На листе '" + nam + "' не найдено окончание месяца '" 
			ask = ask + "--------------------'" + chr(10) + chr(12)
			ask = ask + "OK - перейти к следующему листу, Отмена - прекратить"
			answ = MsgBox(ask, MB_OKCANCEL)
			if answ = IDOK then goto nextsn
			exit for
		end if
		
		' все в порядке - вставляем новый месяц
		InsCopyPrevMonth sn, sh
		cell = sheet.getCellByPosition(1, 0)	' (col, row)
		cell.string = hdr			' делаем новый заголовок
		
		' переносим показания счетчиков
		dim dst as object, r as integer
		for r = 2 to sh - 2
			cell =  sheet.getCellByPosition(3, r)	' (col, row)
			if cell.type = com.sun.star.table.CellContentType.VALUE then
				dst = sheet.getCellByPosition(2, r)	' (col, row)
				dst.value = cell.value
			end if
		next r
		
		' переносим новые значения тарифов и платежей типа E1234.56
		dim col as integer, nc as integer, nv as double
		for r = 2 to sh - 2
			for col = 6 to 16
				cell = sheet.getCellByPosition(col, r)	' (col, row)
				if IsReplaceCell(nc, nv, cell) then		' в ячейке - описание замены?
					cell.string = "." + cell.string
					cell = sheet.getCellByPosition(nc, r)	' (col, row)
					cell.value = nv
				end if
			next col
		next r
		
nextsn:
	next sn
	
end sub


'----------------------------------------------------------------------

private Sub Main

	'CountMonthHeight

	' Получение текущей даты и извлечение из нее года и месяца
	'da = Now
	'y = Year(da)
	'm = Month(da)

	'dim curmony as string
	'curmony = CreateCurMonthYear


	'dim h = CountMonthHeight(0)
	'dim action as integer
	'dim style
	'style = MB_YESNO
	'if h = -1 then
	'	action = MsgBox("Высота не найдена", style)
	'	if action = IDYES then
	'		MsgBox "yes"
	'	elseif action = IDNO then
	'		MsgBox "no"
	'	else
	'		MsgBox "neither yes nor no"
	'	endif
	'end if

	' Хорошая штука, работает
	' InsCopyPrevMonth 0, 14

	Dim Doc as object
	dim sheet as object
	dim cell as object
	dim cnt as integer
	
	Doc = ThisComponent
	cnt = Doc.Sheets.Count
	Sheet = Doc.Sheets(0)
	
	cell = sheet.getCellByPosition(4, 2)	' (col, row)
	dim val as double, stri as string, formu as string
	'val = cell.value
	'str = cell.string
	'formu = cell.formula
	
	Select Case Cell.Type
	Case com.sun.star.table.CellContentType.EMPTY
	  MsgBox "Content: Empty"
	Case com.sun.star.table.CellContentType.VALUE
	  val = Cell.getValue()
	  MsgBox "Content: Value"
	Case com.sun.star.table.CellContentType.TEXT
	  stri = Cell.getString()
	  MsgBox "Content: Text"
	Case com.sun.star.table.CellContentType.FORMULA
	  formu = Cell.getFormula()
	  val = Cell.getValue()
	  MsgBox "Content: Formula"
	End Select	
	
	dim cena as string 
	cena = cell.name


	' Вставка одной строки в позиции 0
	'sheet.rows.insertByIndex(0, 1)

	'TraverceThroughParagraphes

	'dim doc as object
	'doc = ThisComponent		' Типа это добывает весь документ
	'dim url as string
	'url = doc.URL
	
	'dim sf as long
	'sf = com.sun.star.frame.FrameSearchFlag.CREATE
	
	'dim da as Date
	'dim y as integer, m as integer
	
	'dim StyleFamilies as Object
	'dim CellStyles as Object
	'dim CellStyle as Object
	'dim i as Integer
	
	'StyleFamilies = Doc.StyleFamilies
	'CellStyles = StyleFamilies.getByName("CellStyles")
	
	'for i = 0 to CellStyles.Count - 1
	'	CellStyle = CellStyles(i)
	'	MsgBox CellStyle.Name
	'Next i
	
	
	
End Sub


'-------------------------------------------------------
' процедура чисто из руководства посмотреть
private sub TraverceThroughParagraphes
	Dim Doc As Object
	Dim Enum As Object
	Dim TextElement As Object
	
	' Create document object
	Doc = ThisComponent
	' Create enumeration object
	Enum = Doc.Text.createEnumeration		' Для таблицы xls не работает
	' loop over all text elements
	
	While Enum.hasMoreElements
	  TextElement = Enum.nextElement
	
	  If TextElement.supportsService("com.sun.star.text.TextTable") Then
	    MsgBox "The current block contains a table."
	  End If
	
	  If TextElement.supportsService("com.sun.star.text.Paragraph") Then
	    MsgBox "The current block contains a paragraph."
	  End If
	
	Wend
end sub


