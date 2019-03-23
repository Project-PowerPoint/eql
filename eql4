Function EQL5_GetElementValue(Data As String, ElementName As String, Optional Char1 As String = """", Optional Char2 As String = """", Optional Char3 As String = ":", Optional Char4 As String = """", Optional Char5 As String = """", Optional Char6 As String = ";", Optional OnError As String = "undefined")
	On Error GoTo handler
	Dim s As Integer, l As Integer
	s = InStr(1, Data, Char1 & ElementName & Char2 & Char3 & Char4, vbTextCompare)
	l = InStr(s, Data, Char5 & Char6, vbTextCompare) - s
	EQL5_GetElementValue = Mid(Data, s + Len(Char1 & ElementName & Char2 & Char3) + Len(Char4), l - Len(Char1 & ElementName & Char2 & Char3) - Len(Char4))
	Exit Function
	handler: EQL5_GetElementValue = "undefined"
End Function
