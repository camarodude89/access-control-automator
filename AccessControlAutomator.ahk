#Persistent
#SingleInstance

;;;;;;;;;;;
;Constants;
;;;;;;;;;;;

global ALL_BUTTON := "Button2"
global SEARCH_BUTTON := "Button1"
global FIRST_NAME_TEXTBOX := "Edit1"
global LAST_NAME_TEXTBOX := "Edit3"
global SEARCH_RESULTS_COUNT := "Static6"

global LogFile := FileOpen(A_WorkingDir "\log.txt", "a")

openCardholderWindow()
{
	if (!WinExist("Cardholder"))
	{
		if (WinExist("P2000 - Mehler"))
		{
			WinActivate, P2000 - Mehler
			Send {Alt}ac
			Sleep 300
		}
		else
		{
			ExitApp
		}
	}
	else
	{
		WinActivate, Cardholder
		Sleep 300
	}
}

searchEmployee(FirstName, LastName)
{
	WinWaitActive, Cardholder
	Sleep 500
	ControlClick, %ALL_BUTTON%, Cardholder,,L
	Sleep 300
	ControlClick, %SEARCH_BUTTON%, Cardholder,,L
	WinWaitActive, Database Search
	ControlSend, %FIRST_NAME_TEXTBOX%, {Ctrl+a}%FirstName%, Database Search
	ControlSend, %LAST_NAME_TEXTBOX%, {Ctrl+a}%LastName%, Database Search
	Send {Enter}
	WinWaitActive, Cardholder
	ControlGetText, SearchResultsCount, %SEARCH_RESULTS_COUNT%, Cardholder
	SearchResultsCount += 0
		
	if (SearchResultsCount = 1)
	{
		return true
	}
	else
	{
		return false
	}
}

addEmployee(FirstName, MiddleName, LastName)
{
	Sleep 500
	ControlClick, %EMPLOYEE_ADD_BUTTON%, Cardholder,,L
	Send %FirstName%{Tab}
	
	if (MiddleName)
	{
		Send %MiddleName%{Tab}
	}
	else
	{
		Send {Tab}
	}
	
	Send %LastName%
	Send +{Tab 6}{Enter}
	Sleep 500
	Send +{Tab 2}
	LogFile.Write(A_Now . " - " . FirstName . " " . LastName . " added as employee`r`n")
}

addBadge(UserID)
{
	Sleep 500
	Send +{Tab 8}{Enter}
	WinWaitActive, Badge
	Send %UserID%{Enter}
	LogFile.Write(A_Now . " - " . UserID . " added as badge`r`n")
	Sleep 500
}

processCSV(CSVFile)
{
	Loop, read, % CSVFile
	{
		if (A_Index < 2)
		{
			continue
		}
		RecordArray := StrSplit(A_LoopReadLine, ",")
		FirstName := RecordArray[1]
		MiddleName := RecordArray[2]
		LastName := RecordArray[3]
		UserID := RecordArray[4]
		
		WinActivate, Cardholder
		
		searchResultsStatus := searchEmployee(FirstName, LastName)
		
		if (!searchResultsStatus)
		{
			LogFile.Write(A_Now . " - " . FirstName . " " . LastName . " not found`r`n")
			addEmployee(FirstName, MiddleName, LastName)
			addBadge(UserID)
		}
		else
		{
			LogFile.Write(A_Now . " - " . FirstName . " " . LastName . " already exists`r`n")
			Sleep 300
			Send +{Tab 2}{HOME}
			addBadge(UserID)
		}
	}
}

openCardholderWindow()

processCSV("C:\Users\amiller\Desktop\AutoHotkey_1.1.30.01\Projects\AccessControlAutomator\Paycom_Parsed_Export.csv")
LogFile.close()
MsgBox "Record entry for P2000 complete!!! Rock on!!!"
ExitApp
#p::Pause
