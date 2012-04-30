Function FormatSQL(strSQL)
   Dim strNewSQL 'As String
   Dim strTemp   'As String
   Dim intCount  'As Integer
   Dim arrToken  'As String
   Dim blnWhen
   Dim blnParentheses

   strNewSQL = UCase(strSQL)
   strTemp = ""
   intCount = 0
   arrToken = Split(strNewSQL, " ")

   For intCount = LBound(arrToken) To UBound(arrToken)
      If Trim(arrToken(intCount)) = "" Then
         arrToken(intCount) = Trim(arrToken(intCount))
      Else
         arrToken(intCount) = Trim(arrToken(intCount)) & " "
      End If
   Next

   strTemp = Join(arrToken, "")

   intCount = 1

   strNewSQL = ""
   blnWhen = False
   blnParentheses = False
   Do Until intCount > Len(strTemp)
	   Select Case True
	   	Case Mid(strTemp, intCount, 7) = "SELECT " And intCount = 1
	   		strNewSQL = strNewSQL & vbCrlf & Mid(strTemp, intCount, 7) & vbCrlf & "<br/>&nbsp;&nbsp;"
   			intCount = intCount + 7
	   	Case Mid(strTemp, intCount, 7) = "SELECT "
	   		strNewSQL = strNewSQL & vbCrlf & "<br/><br/>" & Mid(strTemp, intCount, 7) & vbCrlf & "<br/>  "
   			intCount = intCount + 7
	   	Case Mid(strTemp, intCount, 7) = "INSERT "
	   		strNewSQL = strNewSQL & vbCrlf & "<br/><br/>" & Mid(strTemp, intCount, 7) & vbCrlf & "<br/>"
   			intCount = intCount + 7
	   	Case Mid(strTemp, intCount, 7) = "FETCH "
	   		strNewSQL = strNewSQL & vbCrlf & "<br/><br/>" & Mid(strTemp, intCount, 6) & vbCrlf & "<br/>"
   			intCount = intCount + 6
	   	Case Mid(strTemp, intCount, 5) = "CASE "
	   		strNewSQL = strNewSQL & vbCrlf & "<br/>" & Mid(strTemp, intCount, 5)
   			intCount = intCount + 5
	   	Case Mid(strTemp, intCount, 6) = " WHEN "
	   		strNewSQL = strNewSQL & vbCrlf & "<br/>" & Mid(strTemp, intCount, 6)
	   		blnWhen = True
   			intCount = intCount + 6
	   	Case Mid(strTemp, intCount, 6) = " ELSE "
	   		strNewSQL = strNewSQL & vbCrlf & "<br/>" & Mid(strTemp, intCount, 6)
   			intCount = intCount + 6
	   	Case Mid(strTemp, intCount, 4) = " END"
	   		strNewSQL = strNewSQL & Mid(strTemp, intCount, 4)
	   		blnWhen = False
   			intCount = intCount + 4
	   	Case Mid(strTemp, intCount, 5) = "INTO "
	   		strNewSQL = strNewSQL & vbCrlf & vbCrlf & "<br/><br/>&nbsp;" & Mid(strTemp, intCount, 5)
   			intCount = intCount + 5
	   	Case Mid(strTemp, intCount, 5) = "FROM "
	   		strNewSQL = strNewSQL & vbCrlf & vbCrlf & "<br/><br/>&nbsp;" & Mid(strTemp, intCount, 5)
   			intCount = intCount + 5
	   	Case Mid(strTemp, intCount, 11) = "INNER JOIN "
	   		strNewSQL = strNewSQL & vbCrlf & vbCrlf & "<br/><br/>" & Mid(strTemp, intCount, 11)
   			intCount = intCount + 11
	   	Case Mid(strTemp, intCount, 5) = "JOIN "
	   		strNewSQL = strNewSQL & vbCrlf & vbCrlf & "<br/><br/>&nbsp;" & Mid(strTemp, intCount, 5)
   			intCount = intCount + 5
	   	Case Mid(strTemp, intCount, 9) = "LEFT JOIN"
	   		strNewSQL = strNewSQL & vbCrlf & vbCrlf & "<br/><br/>&nbsp;" & Mid(strTemp, intCount, 9)
   			intCount = intCount + 9
	   	Case Mid(strTemp, intCount, 15) = "LEFT OUTER JOIN"
	   		strNewSQL = strNewSQL & vbCrlf & vbCrlf & "<br/><br/>&nbsp;" & Mid(strTemp, intCount, 15)
   			intCount = intCount + 15
	   	Case Mid(strTemp, intCount, 15) = "FULL OUTER JOIN"
	   		strNewSQL = strNewSQL & vbCrlf & vbCrlf & "<br/><br/>&nbsp;" & Mid(strTemp, intCount, 15)
   			intCount = intCount + 15
	   	Case Mid(strTemp, intCount, 6) = "UNION "
	   		strNewSQL = strNewSQL & vbCrlf & vbCrlf & "<br/><br/>" & Mid(strTemp, intCount, 6)
   			intCount = intCount + 6
	   	Case Mid(strTemp, intCount, 5) = " AND " and Not(blnWhen)
	   		strNewSQL = strNewSQL & vbCrlf & "<br/>&nbsp;" & Mid(strTemp, intCount, 5)
   			intCount = intCount + 5
	   	Case Mid(strTemp, intCount, 4) = " OR "
	   		strNewSQL = strNewSQL & vbCrlf & "<br/>" & Mid(strTemp, intCount, 4)
   			intCount = intCount + 4
	   	Case Mid(strTemp, intCount, 6) = "WHERE "
	   		strNewSQL = strNewSQL & vbCrlf & vbCrlf & "<br/><br/>" & Mid(strTemp, intCount, 6)
   			intCount = intCount + 6
	   	Case Mid(strTemp, intCount, 4) = "IN ("
	   		strNewSQL = strNewSQL & Mid(strTemp, intCount, 4)
	   		blnParentheses = True
   			intCount = intCount + 4
	   	Case Mid(strTemp, intCount, 7) = "SUBSTR("
	   		strNewSQL = strNewSQL & Mid(strTemp, intCount, 7)
	   		blnParentheses = True
   			intCount = intCount + 7
	   	Case Mid(strTemp, intCount, 6) = "VALUE("
	   		strNewSQL = strNewSQL & Mid(strTemp, intCount, 6)
	   		blnParentheses = True
   			intCount = intCount + 6
	   	Case Mid(strTemp, intCount, 9) = "COALESCE("
	   		strNewSQL = strNewSQL & Mid(strTemp, intCount, 9)
	   		blnParentheses = True
   			intCount = intCount + 9
	   	Case Mid(strTemp, intCount, 6) = "RTRIM("
	   		strNewSQL = strNewSQL & Mid(strTemp, intCount, 6)
	   		blnParentheses = True
   			intCount = intCount + 6
	   	Case Mid(strTemp, intCount, 6) = "LTRIM("
	   		strNewSQL = strNewSQL & Mid(strTemp, intCount, 6)
	   		blnParentheses = True
   			intCount = intCount + 6
	   	Case Mid(strTemp, intCount, 13) = "MULTIPLY_ALT("
	   		strNewSQL = strNewSQL & Mid(strTemp, intCount, 13)
	   		blnParentheses = True
   			intCount = intCount + 13
	   	Case Mid(strTemp, intCount, 7) = "NULLIF("
	   		strNewSQL = strNewSQL & Mid(strTemp, intCount, 7)
	   		blnParentheses = True
   			intCount = intCount + 7
	   	Case Mid(strTemp, intCount, 8) = "REPLACE("
	   		strNewSQL = strNewSQL & Mid(strTemp, intCount, 8)
	   		blnParentheses = True
   			intCount = intCount + 8
	   	Case Mid(strTemp, intCount, 8) = "REPLACE("
	   		strNewSQL = strNewSQL & Mid(strTemp, intCount, 8)
	   		blnParentheses = True
   			intCount = intCount + 8
	   	Case Mid(strTemp, intCount, 1) = ")"
	   		strNewSQL = strNewSQL & Mid(strTemp, intCount, 1)
	   		blnParentheses = False
   			intCount = intCount + 1
	   	Case Mid(strTemp, intCount, 1) = "," and Not(blnParentheses)
	   		strNewSQL = strNewSQL & vbCrlf & "<br/>&nbsp;" & Mid(strTemp, intCount, 1)
	   		blnParentheses = False
   			intCount = intCount + 1
	   	Case Mid(strTemp, intCount, 5) = " SET "
	   		strNewSQL = strNewSQL & vbCrlf & vbCrlf & "<br/><br/>" & Mid(strTemp, intCount, 4)
   			intCount = intCount + 4
	   	Case Mid(strTemp, intCount, 6) = " VALUES"
	   		strNewSQL = strNewSQL & vbCrlf & vbCrlf & "<br/><br/>" & Mid(strTemp, intCount, 6) & vbCrlf & "<br/>"
   			intCount = intCount + 6
	   	Case Mid(strTemp, intCount, 4) = " ON "
	   		strNewSQL = strNewSQL & vbCrlf & "<br/>&nbsp;&nbsp;" & Mid(strTemp, intCount, 4)
   			intCount = intCount + 4
	   	Case Mid(strTemp, intCount, 8) = "GROUP BY"
	   		strNewSQL = strNewSQL & vbCrlf & vbCrlf & "<br/><br/>" & Mid(strTemp, intCount, 8)
   			intCount = intCount + 8
	   	Case Mid(strTemp, intCount, 8) = "ORDER BY"
	   		strNewSQL = strNewSQL & vbCrlf & vbCrlf & "<br/><br/>" & Mid(strTemp, intCount, 8)
   			intCount = intCount + 8
	   	Case Else
	   		strNewSQL = strNewSQL & Mid(strTemp, intCount, 1)
   			intCount = intCount + 1
   	End Select
	Loop
	strNewSQL = strNewSQL & "<br/><br/>" & vbCrlf & vbCrlf

   FormatSQL = strNewSQL
End Function