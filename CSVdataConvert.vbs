Option Explicit
	
Dim WshShell
  Set WshShell               = WScript.CreateObject("WScript.Shell")
Dim fs,f
  Set fs                     = CreateObject("Scripting.FileSystemObject")
Dim ObjFolder, FolderFiles
Dim Arg, FileName
Dim ReqArr, ResArr,StepArr,TestArr
	
Dim objExcel, Sheet, objSheet, objCells, strExcelPath, r, UpRow
Dim objProgressMsg
Dim PopUpTitle
  PopUpTitle = "CSV Converter"
Dim ResultArray, LundiArr, MardiArr, MercrediArr, JeudiArr, VendrediArr, SamediArr, DimancheArr, TotalEvent, ErrTotalNum
	LundiArr	= Array()
	MardiArr	= Array()
	MercrediArr	= Array()
	JeudiArr	= Array()
	VendrediArr	= Array()
	SamediArr	= Array()
	DimancheArr	= Array()
	ErrTotalNum	= Array()
Dim percLun, percMar, percMer, percJeu, percVen, percSam, percDim

	
If WScript.Arguments.Count > 0 Then
    For Each Arg in Wscript.Arguments
        Arg =  Trim(Arg)
    If InStr(Arg,".") Then
		If UCase(fs.GetExtensionName(Arg)) = "CSV" Then
			'Set f = fs.OpenTextFile(Arg, 1)
			FileName = Trim(Replace(Split(Arg, "\")(Ubound(Split(Arg, "\"))),".csv",""))
			Call ParseCSV(Arg)
		End If
    End If
    Next
Else

		  MsgBox "Please Drag&Drop your .csv files onto the script"
End If



''----------------------------------------------------------------------------------''
Sub ParseCSV(InputFile)

	Dim strLine, objInputFile
	Dim objFSO
	Dim SkipFirstLine
		SkipFirstLine = 0
	Dim DayNum, ErrNum, Found
	Dim WeekArr,EventElement,CurrEventNum,AlreadyGotEvent,ExitVar,ErrCheckElement,ErrCElement
		WeekArr = Array("0","LundiArr","MardiArr","MercrediArr","JeudiArr","VendrediArr","SamediArr","DimancheArr")

		
		Set objFSO = CreateObject("Scripting.FileSystemObject") 
		Set objInputFile = objFSO.OpenTextFile(InputFile, 1)  '------- Input file


	Do While objInputFile.AtEndOfStream = False			'------- Line by line execution until the end of the file
		strLine = objInputFile.ReadLine
		If SkipFirstLine=0 Then
			SkipFirstLine = 1
		Else
		
			DayNum = Weekday(Split(Split(strLine,"""")(1)," ")(0),2)
			ErrNum = Split(strLine,"""")(3)

			If DayNum = 1 Then
				ReDim Preserve LundiArr(UBound(LundiArr) + 1)
				LundiArr(UBound(LundiArr)) = ErrNum
			ElseIf DayNum = 2 Then
				ReDim Preserve MardiArr(UBound(MardiArr) + 1)
				MardiArr(UBound(MardiArr)) = ErrNum		
			ElseIf DayNum = 3 Then
				ReDim Preserve MercrediArr(UBound(MercrediArr) + 1)
				MercrediArr(UBound(MercrediArr)) = ErrNum		
			ElseIf DayNum = 4 Then
				ReDim Preserve JeudiArr(UBound(JeudiArr) + 1)
				JeudiArr(UBound(JeudiArr)) = ErrNum		
			ElseIf DayNum = 5 Then
				ReDim Preserve VendrediArr(UBound(VendrediArr) + 1)
				VendrediArr(UBound(VendrediArr)) = ErrNum		
			ElseIf DayNum = 6 Then
				ReDim Preserve SamediArr(UBound(SamediArr) + 1)
				SamediArr(UBound(SamediArr)) = ErrNum		
			ElseIf DayNum = 7 Then
				ReDim Preserve DimancheArr(UBound(DimancheArr) + 1)
				DimancheArr(UBound(DimancheArr)) = ErrNum		
			End If
			
			Found = 0
			For Each ErrCheckElement in ErrTotalNum
				If StrComp(ErrCheckElement(0),ErrNum)=0 Then
					ErrCheckElement(1) = ErrCheckElement(1) + 1
					'MsgBox "/" & ErrCheckElement(1)  ''marche pas
					Found = 1
				End If
			Next
			
			If Found = 0 Then
				ReDim Preserve ErrTotalNum(UBound(ErrTotalNum) + 1)
				ErrTotalNum(UBound(ErrTotalNum)) = Array(ErrNum,1,0)	
			End If
			''__ErrTotalNum : (ErrNum,#)(ErrNum2,#)... -> stat par Erreur
		End If
	Loop

	''___ Total # of event and % per day ___''
	TotalEvent = UBound(LundiArr)+UBound(MardiArr)+UBound(MercrediArr)+UBound(JeudiArr)+UBound(VendrediArr)+UBound(SamediArr)+UBound(DimancheArr)
	percLun = Round((UBound(LundiArr)*100)/TotalEvent,2)
	percMar = Round((UBound(MardiArr)*100)/TotalEvent,2)
	percMer = Round((UBound(MercrediArr)*100)/TotalEvent,2)
	percJeu = Round((UBound(JeudiArr)*100)/TotalEvent,2)
	percVen = Round((UBound(VendrediArr)*100)/TotalEvent,2)
	percSam = Round((UBound(SamediArr)*100)/TotalEvent,2)
	percDim = Round((UBound(DimancheArr)*100)/TotalEvent,2)
	
	For Each ErrCheckElement in ErrTotalNum
		ErrCheckElement(2)=(ErrCheckElement(1)*100)/TotalEvent
		''MsgBox ErrCheckElement(1) & "  -  " & ErrCheckElement(2) & "%"	''marche pas 
	Next
	
MsgBox "Nombre Total d'evenements : " & TotalEvent & Chr(13) & "Le Lundi : " & percLun & "%" & Chr(13) & "Le Mardi : " & percMar & "%" & Chr(13) & "Le Mercredi : " & percMer & "%" & Chr(13) & "Le Jeudi : " & percJeu & "%" & Chr(13) & "Le Vendredi : " & percVen & "%" & Chr(13) & "Le Samedi : " & percSam & "%" & Chr(13) & "Le Dimanche : " & percDim & "%" & Chr(13)


	
End Sub