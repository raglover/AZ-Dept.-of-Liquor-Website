<%@LANGUAGE="VBSCRIPT"%>
<% Option Explicit %>
<%
'Standalone captcha ASP CODE
'http://www.hotdreamweaver.com
     
Dim cw, ch, LDistortNum, secureCode, rowNum, i, j, RangeSize,tmpNum,DistortNum,clmNum, ColX, RowY
Dim CharWidth, CodeLength, Distort, DistortEx, Noise, TClr, BClr, NClr, LeftMargin, BottomMargin, UseRandomColors
Dim newBitmap(1000,1000)
Dim vDistort(100)
LDistortNum = 0
Const CharHeight = 13

if (Request.QueryString("width")="") Then cw = 106 Else cw = CInt(Request.QueryString("width")) 

While ( (cw-2) Mod 4 <> 0)
  cw = cw - 1
WEnd
  
if (Request.QueryString("height")="") Then ch = 40 Else ch = CInt(Request.QueryString("height")) 
if (Request.QueryString("letter_count")="") Then CodeLength = 6 Else CodeLength = CInt(Request.QueryString("letter_count"))
if (Request.QueryString("min_size")="") Then CharWidth = 14 Else CharWidth = CInt(Request.QueryString("min_size"))
If (Request.QueryString("noise")="") Then Noise = 2 Else Noise = CInt(Request.QueryString("noise")) 

If (Request.QueryString("distort")="0") Then
  Distort = False
  DistortEx = False   
Else
  Distort = True 
  DistortEx = True   
End If

If (Request.QueryString("tcolor")="") Then TClr = "990000" Else TClr = InverseColor(Request.QueryString("tcolor"))
If (Request.QueryString("bcolor")="") Then BClr = "FFFFFF" Else BClr = InverseColor(Request.QueryString("bcolor"))
If (Request.QueryString("ncolor")="") Then NClr = "990000" Else NClr = InverseColor(Request.QueryString("ncolor"))
If (Request.QueryString("rcolor")="1") Then UseRandomColors = True Else UseRandomColors = False

LeftMargin = CInt(cw/2 - (91*CodeLength*CharWidth/(8*11.6))/2)
BottomMargin = CInt(ch/2 - 14/2)
If (LeftMargin < 0) Then LeftMargin = 0
If (BottomMargin < 0) Then BottomMargin = 0


'End editable consts
Dim BmpHeader 
BmpHeader = "424D8C150000000000003600000028000000" & Hex(cw) & "000000" & Hex(ch) & "000000010018000000000056150000120B0000120B00000000000000000000"
Const BmpEndLine = "0000"

Function InverseColor (color)
  InverseColor = UCase(Right(color,2) & Mid(color,3,2) & Left(color,2))
End Function

Function RandomColor
  Dim c
  c = "0123456789ABCDEF"
  RandomColor = Mid(c,Random(1,16),1)&Mid(c,Random(1,16),1)&Mid(c,Random(1,16),1)&Mid(c,Random(1,16),1)&Mid(c,Random(1,16),1)&Mid(c,Random(1,16),1)
End Function

Sub IHex(iRow,iColumn,strHex,iRepeat)
    Dim x
	for x=0 to (iRepeat-1)
		newBitmap(iRow,iColumn+x) = strHex
	next
End Sub

Function Random(valMin,valMax)
    Randomize(timer)
    RangeSize = ((valMax - valMin) + 1)
    Random = Int((RangeSize * Rnd()) + 1)
End Function

Sub AddNoise()
    Dim x, y
    for x=0 to CInt(cw*28/86)
        for y=0 to CInt(ch*7/21)
            If UseRandomColors Then NClr = RandomColor
            if (Random(1,Noise) = 1) Then
                ColX = (x*3) + Random(1,3)
                RowY = (y*3) + Random(1,3)
                IHex RowY,ColX,NClr,1
            End If    
            If UseRandomColors Then NClr = RandomColor
            If (Random(1,Noise) = 1) Then    
                ColX = (x*3) + Random(1,3)
                RowY = (y*3) + Random(1,3)
                IHex RowY,ColX,NClr,1
            End If    
        next
    next
End Sub

Sub WriteCanvas(valChar,iNumPart,iRow,iColumn)
	select case iNumPart
		case 1
			select case valChar
				case 0
					IHex iRow,iColumn+2,TClr,4
				case 1
					IHex iRow,iColumn+3,TClr,2
				case 2
					IHex iRow,iColumn+2,TClr,4
				case 3
					IHex iRow,iColumn+2,TClr,3
				case 4
					IHex iRow,iColumn+5,TClr,2
				case 5
					IHex iRow,iColumn+1,TClr,6
				case 6
					IHex iRow,iColumn+2,TClr,4
				case 7
					IHex iRow,iColumn,TClr,8
				case 8
					IHex iRow,iColumn+2,TClr,4
				case 9
					IHex iRow,iColumn+2,TClr,4
			end select
		case 2
			select case valChar
				case 0
					IHex iRow,iColumn+1,TClr,6
				case 1
					IHex iRow,iColumn+2,TClr,3
				case 2
					IHex iRow,iColumn+1,TClr,6
				case 3
					IHex iRow,iColumn+1,TClr,6
				case 4
					IHex iRow,iColumn+4,TClr,3
				case 5
					IHex iRow,iColumn+1,TClr,6
				case 6
					IHex iRow,iColumn+1,TClr,6
				case 7
					IHex iRow,iColumn,TClr,8
				case 8
					IHex iRow,iColumn+1,TClr,6
				case 9
					IHex iRow,iColumn+1,TClr,6
			end select
		case 3
			select case valChar
				case 0
					IHex iRow,iColumn,TClr,3
					IHex iRow,iColumn+5,TClr,3
				case 1
					IHex iRow,iColumn+1,TClr,4
				case 2
					IHex iRow,iColumn,TClr,3
					IHex iRow,iColumn+5,TClr,3
				case 3
					IHex iRow,iColumn,TClr,2
					IHex iRow,iColumn+5,TClr,2
				case 4
					IHex iRow,iColumn+4,TClr,3
				case 5
					IHex iRow,iColumn+1,TClr,2
				case 6
					IHex iRow,iColumn+1,TClr,2
					IHex iRow,iColumn+6,TClr,2
				case 7
					IHex iRow,iColumn+6,TClr,1
				case 8
					IHex iRow,iColumn,TClr,2
					IHex iRow,iColumn+6,TClr,2
				case 9
					IHex iRow,iColumn,TClr,2
					IHex iRow,iColumn+6,TClr,2
			end select
		case 4
			select case valChar
				case 0
					IHex iRow,iColumn,TClr,2
					IHex iRow,iColumn+6,TClr,2
				case 1
					IHex iRow,iColumn,TClr,2
					IHex iRow,iColumn+3,TClr,2
				case 2
					IHex iRow,iColumn,TClr,2
					IHex iRow,iColumn+6,TClr,2
				case 3
					IHex iRow,iColumn+5,TClr,2
				case 4
					IHex iRow,iColumn+3,TClr,4
				case 5
					IHex iRow,iColumn,TClr,2
				case 6
					IHex iRow,iColumn,TClr,2
				case 7
					IHex iRow,iColumn+5,TClr,2
				case 8
					IHex iRow,iColumn,TClr,2
					IHex iRow,iColumn+6,TClr,2
				case 9
					IHex iRow,iColumn,TClr,2
					IHex iRow,iColumn+6,TClr,2
			end select
		case 5
			select case valChar
				case 0
					IHex iRow,iColumn,TClr,2
					IHex iRow,iColumn+6,TClr,2
				case 1
					IHex iRow,iColumn,TClr,1
					IHex iRow,iColumn+3,TClr,2
				case 2
					IHex iRow,iColumn+6,TClr,2
				case 3
					IHex iRow,iColumn+5,TClr,2
				case 4
					IHex iRow,iColumn+2,TClr,2
					IHex iRow,iColumn+5,TClr,2
				case 5
					IHex iRow,iColumn,TClr,2
					IHex iRow,iColumn+3,TClr,3
				case 6
					IHex iRow,iColumn,TClr,2
				case 7
					IHex iRow,iColumn+4,TClr,2
				case 8
					IHex iRow,iColumn,TClr,2
					IHex iRow,iColumn+6,TClr,2
				case 9
					IHex iRow,iColumn,TClr,2
					IHex iRow,iColumn+6,TClr,2
			end select
		case 6
			select case valChar
				case 0
					IHex iRow,iColumn,TClr,2
					IHex iRow,iColumn+6,TClr,2
				case 1
					IHex iRow,iColumn+3,TClr,2
				case 2
					IHex iRow,iColumn+6,TClr,2
				case 3
					IHex iRow,iColumn+3,TClr,3
				case 4
					IHex iRow,iColumn+2,TClr,2
					IHex iRow,iColumn+5,TClr,2
				case 5
					IHex iRow,iColumn,TClr,7
				case 6
					IHex iRow,iColumn,TClr,2
					IHex iRow,iColumn+3,TClr,3
				case 7
					IHex iRow,iColumn+4,TClr,2
				case 8
					IHex iRow,iColumn+1,TClr,6
				case 9
					IHex iRow,iColumn,TClr,2
					IHex iRow,iColumn+5,TClr,3
			end select
		case 7
			select case valChar
				case 0
					IHex iRow,iColumn,TClr,2
					IHex iRow,iColumn+6,TClr,2
				case 1
					IHex iRow,iColumn+3,TClr,2
				case 2
					IHex iRow,iColumn+5,TClr,2
				case 3
					IHex iRow,iColumn+3,TClr,4
				case 4
					IHex iRow,iColumn+1,TClr,2
					IHex iRow,iColumn+5,TClr,2
				case 5
					IHex iRow,iColumn,TClr,2
					IHex iRow,iColumn+6,TClr,2
				case 6
					IHex iRow,iColumn,TClr,7
				case 7
					IHex iRow,iColumn+3,TClr,2
				case 8
					IHex iRow,iColumn+1,TClr,6
				case 9
					IHex iRow,iColumn+1,TClr,7
			end select
		case 8
			select case valChar
				case 0
					IHex iRow,iColumn,TClr,2
					IHex iRow,iColumn+6,TClr,2
				case 1
					IHex iRow,iColumn+3,TClr,2
				case 2
					IHex iRow,iColumn+4,TClr,2
				case 3
					IHex iRow,iColumn+6,TClr,2
				case 4
					IHex iRow,iColumn,TClr,2
					IHex iRow,iColumn+5,TClr,2
				case 5
					IHex iRow,iColumn+6,TClr,2
				case 6
					IHex iRow,iColumn,TClr,3
					IHex iRow,iColumn+6,TClr,2
				case 7
					IHex iRow,iColumn+3,TClr,2
				case 8
					IHex iRow,iColumn,TClr,2
					IHex iRow,iColumn+6,TClr,2
				case 9
					IHex iRow,iColumn+2,TClr,3
					IHex iRow,iColumn+6,TClr,2
			end select
		case 9
			select case valChar
				case 0
					IHex iRow,iColumn,TClr,2
					IHex iRow,iColumn+6,TClr,2
				case 1
					IHex iRow,iColumn+3,TClr,2
				case 2
					IHex iRow,iColumn+3,TClr,2
				case 3
					IHex iRow,iColumn+6,TClr,2
				case 4
					IHex iRow,iColumn,TClr,9
				case 5
					IHex iRow,iColumn+6,TClr,2
				case 6
					IHex iRow,iColumn,TClr,2
					IHex iRow,iColumn+6,TClr,2
				case 7
					IHex iRow,iColumn+3,TClr,2
				case 8
					IHex iRow,iColumn,TClr,2
					IHex iRow,iColumn+6,TClr,2
				case 9
					IHex iRow,iColumn+6,TClr,2
			end select
		case 10
			select case valChar
				case 0
					IHex iRow,iColumn,TClr,2
					IHex iRow,iColumn+6,TClr,2
				case 1
					IHex iRow,iColumn+3,TClr,2
				case 2
					IHex iRow,iColumn+2,TClr,2
				case 3
					IHex iRow,iColumn,TClr,2
					IHex iRow,iColumn+6,TClr,2
				case 4
					IHex iRow,iColumn,TClr,9
				case 5
					IHex iRow,iColumn,TClr,2
					IHex iRow,iColumn+6,TClr,2
				case 6
					IHex iRow,iColumn,TClr,2
					IHex iRow,iColumn+6,TClr,2
				case 7
					IHex iRow,iColumn+3,TClr,2
				case 8
					IHex iRow,iColumn,TClr,2
					IHex iRow,iColumn+6,TClr,2
				case 9
					IHex iRow,iColumn+6,TClr,2
			end select
		case 11
			select case valChar
				case 0
					IHex iRow,iColumn,TClr,3
					IHex iRow,iColumn+5,TClr,3
				case 1
					IHex iRow,iColumn+3,TClr,2
				case 2
					IHex iRow,iColumn+1,TClr,2
				case 3
					IHex iRow,iColumn,TClr,2
					IHex iRow,iColumn+6,TClr,2
				case 4
					IHex iRow,iColumn+5,TClr,2
				case 5
					IHex iRow,iColumn,TClr,2
					IHex iRow,iColumn+6,TClr,2
				case 6
					IHex iRow,iColumn+1,TClr,2
					IHex iRow,iColumn+6,TClr,2
				case 7
					IHex iRow,iColumn+2,TClr,2
				case 8
					IHex iRow,iColumn,TClr,2
					IHex iRow,iColumn+6,TClr,2
				case 9
					IHex iRow,iColumn,TClr,2
					IHex iRow,iColumn+5,TClr,2
			end select
		case 12
			select case valChar
				case 0
					IHex iRow,iColumn+1,TClr,6
				case 1
					IHex iRow,iColumn+3,TClr,2
				case 2
					IHex iRow,iColumn,TClr,8
				case 3
					IHex iRow,iColumn+1,TClr,6
				case 4
					IHex iRow,iColumn+5,TClr,2
				case 5
					IHex iRow,iColumn+1,TClr,6
				case 6
					IHex iRow,iColumn+1,TClr,6
				case 7
					IHex iRow,iColumn+2,TClr,2
				case 8
					IHex iRow,iColumn+1,TClr,6
				case 9
					IHex iRow,iColumn+1,TClr,6
			end select
		case 13
			select case valChar
				case 0
					IHex iRow,iColumn+2,TClr,4
				case 1
					IHex iRow,iColumn+3,TClr,2
				case 2
					IHex iRow,iColumn,TClr,8
				case 3
					IHex iRow,iColumn+2,TClr,4
				case 4
					IHex iRow,iColumn+5,TClr,2
				case 5
					IHex iRow,iColumn+2,TClr,4
				case 6
					IHex iRow,iColumn+2,TClr,4
				case 7
					IHex iRow,iColumn+2,TClr,2
				case 8
					IHex iRow,iColumn+2,TClr,4
				case 9
					IHex iRow,iColumn+2,TClr,4
			end select
	end select
End Sub

Function LeftTracking(iNumber)
	select case iNumber
		case 1
			LeftTracking = 2
		case 4
			LeftTracking = 0
		case else
			LeftTracking = 1
	end select
End Function

Function CreateGUID(tmpLength)
  Randomize Timer
  Dim tmpCounter,tmpGUID
  Const strValid = "01234567890"
  For tmpCounter = 1 To tmpLength
    tmpGUID = tmpGUID & Mid(strValid, Int(Rnd(1) * Len(strValid)) + 1, 1)
  Next
  CreateGUID = tmpGUID
End Function

Function GetStartColumn(iNumber,iLine)
	if DistortEx = true then
		DistortNum = (Random(1,3) - 1)
		if DistortNum = 0 then
		    DistortNum = LDistortNum
		end if
		LDistortNum = DistortNum
	else
		DistortNum = 0
	end if
	GetStartColumn =  LeftMargin + ((CharWidth * (iLine-1)) + LeftTracking(iNumber)) + DistortNum
End Function

Sub SendHex(valHex)
    Dim k, strHex
	for k=1 to Len(valHex)
		strHex = "&H" & Mid(valHex,k,2)
		Response.BinaryWrite ChrB(CInt(strHex))
		k=k+1
	next
End Sub

Sub SendClient()
	Response.Buffer = True
	Response.ContentType = "image/bmp"
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

    Dim x, y, tmpHex

    if Noise > 0 Then
        AddNoise()
    End If    
    
	SendHex(BmpHeader)
	for y=1 to ch
		for x=1 to cw
			tmpHex = newBitmap(y,x)
			if tmpHex = vbNullString then
				SendHex(BClr)
			else
				SendHex(tmpHex)
			end if

			if x=cw then
				SendHex(BmpEndLine)
			end if
		next
	next
	
	SendHex(BmpEndLine)
	Response.Flush
End Sub
%>

<%
secureCode = CreateGUID(CodeLength)
Session("HDWCAPTCHA") = secureCode

for i=1 to CharHeight
	rowNum = (ch - (BottomMargin + (i-1)))
	for j=1 to Len(secureCode)
		if (Distort = true) and (i=1) then
			vDistort(j) = (Random(1,6) - 3)
		elseif (i=1) then
			vDistort(j) = 0
		end if
		tmpNum = CInt(Mid(secureCode,j,1))
		clmNum = GetStartColumn(tmpNum,j)
		If UseRandomColors Then TClr = RandomColor
		WriteCanvas tmpNum,i,rowNum+vDistort(j),clmNum
	next
next

SendClient()
%>