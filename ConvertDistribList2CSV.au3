#include <Array.au3>
#include <File.au3>
#include <FileConstants.au3>
#include <MsgBoxConstants.au3>
#include <MsgBoxConstants.au3>
#include <StringConstants.au3>
#include <AutoItConstants.au3>

; This script is intended to parse a file exported from exchange powershell with distribution lists listed like so
;
; dist1 (dist1@domain.com)
; =============
; First Member Name (fmember@domain.com)
; Second Member Name (smember@domain.com)
;
;Script will prompt for the location of the file, will then show the file as an array, once closed it will process
;the file if its in the proper format, into rows for each email address associated with each distribution list for auditing purposes.
;the final output will be as follows for columns in the CSV:
;DISTRO ID, DEPTARTMENT, OWNER, PURPOSE, MEMBER COUND, EMAIL ADDRESS

Local Const $sFilePath = ""
Local Const $sOutputFileName = $sFilePath & ""
Local $iTimer = TimerInit()
Local $sFileRead = ""
Local $hFileOpen

Main()


Func Main()
	$sProgressDot = "."
	ToolTip($sProgressDot,0,0,"[....................] ConvertDistribList Running")
 ; Create a constant variable in Local scope of the message to display in FileOpenDialog.
    Local Const $sMessage = "Choose File to Convert - Hold down Ctrl or Shift to choose multiple files."

    ; Display an open dialog to select a list of file(s).
    Local $sFileOpenDialog = FileOpenDialog($sMessage, @WindowsDir & "\", "Text (*.txt)", BitOR($FD_FILEMUSTEXIST, $FD_MULTISELECT))
    If @error Then
        ; Display the error message.
		;DeBug("No file(s) were selected.")
        ;DeBug("No file(s) were selected.")

        ; Change the working directory (@WorkingDir) back to the location of the script directory as FileOpenDialog sets it to the last accessed folder.
        FileChangeDir(@ScriptDir)
    Else
        ; Change the working directory (@WorkingDir) back to the location of the script directory as FileOpenDialog sets it to the last accessed folder.
        FileChangeDir(@ScriptDir)

        ;Load file as an Array for iteration
        Local $aArray = FileReadToArray($sFileOpenDialog)
		Local $iLineCount = @extended
		If @error Then
			;DeBug("There was an error reading the file. @error: " & @error) ; An error occurred reading the current script file.
		Else
			; Display the list of selected files.
			;DeBug("You chose the following files:" & @CRLF & $sFileOpenDialog)
			;Display the files converted to an Array
			_ArrayDisplay($aArray)
			Local $aCSVRows[1000][6]
			Local $iCSVRow = 0
			Local $iBlankCount = 0
			Local $bIsSeparator = False
			Local $bIsDistro = True
			Local $bIsMember = False
			Local $sCurrentDistro = ""
			Local $sCurrentMember = ""
			For $i = 0 to UBound($aArray)-1
				ToolTip($sProgressDot,0,0,"ConvertDistribList Running")
				;DeBug("$aArray[" & $i & "]:" & $aArray[$i])
				If $aArray[$i] = "" Then
					$bIsDistro = True
					$bIsMember = False
					Local $sCurrentMember = ""
					$sProgressDot = "."
					;DeBug("IsBlank1:" & $iBlankCount)
					ContinueLoop
				ElseIf $aArray[$i] = "=============" Then
					$iBlankCount = 0
					$bIsSeparator = True
					$bIsMember = True
					$bIsDistro = False
					Local $sCurrentMember = ""
					$sProgressDot = ""
					;DeBug("Is ==========:" & $iBlankCount)
					ContinueLoop
				Else
					If UBound($aCSVRows) = $iCSVRow  Then
						ReDim $aCSVRows[UBound($aCSVRows) + 1000][6]
					EndIf
					$iBlankCount = 0
					Select
						Case $bIsDistro = True
							$sCurrentDistro = $aArray[$i]
							$CurrentMember = ""
							$bIsDistro = False
							;DeBug("Case: $bDistro = True: " & $sCurrentDistro)
						Case $bIsDistro = False
							$sCurrentMember = $aArray[$i]
							;DeBug("CaseElse: $bDistro = False: CurMember" & $sCurrentMember)
					EndSelect
					$aCSVRows[$iCSVRow][0] = $sCurrentDistro ;DistroName
					$aCSVRows[$iCSVRow][1] = "-" ; Dept
					$aCSVRows[$iCSVRow][2] = "-" ; Purpose
					$aCSVRows[$iCSVRow][3] = "-" ; Owner
					$aCSVRows[$iCSVRow][4] = "-" ; Purpose
					$aCSVRows[$iCSVRow][5] = $sCurrentMember
					$iCSVRow = $iCSVRow + 1
					If StringLen($sProgressDot) > 100 Then
						$sProgressDot = "."
					Else
						$sProgressDot = $sProgressDot & "."
					EndIf

					If UBound($aCSVRows) < $iCSVRow Then
						ReDim $aCSVRows[UBound($aCSVRows) + 1000][6]
					EndIf
					;DeBug("CaseElse:" & $iBlankCount)
					;DeBug($aCSVRows)
				EndIf
			Next
		EndIf
		_ArrayDisplay($aCSVRows)

	; Create a constant variable in Local scope of the message to display in FileSaveDialog.

    ; Display a save dialog to select a file.
    Local $sFileSaveDialog = FileSaveDialog($sMessage, "", "CSV (*.csv)", $FD_PATHMUSTEXIST)
    If @error Then
        ; Display the error message.
        MsgBox($MB_SYSTEMMODAL, "", "No file was saved.")
    Else
        ; Retrieve the filename from the filepath e.g. Example.au3.
        Local $sFileName = StringTrimLeft($sFileSaveDialog, StringInStr($sFileSaveDialog, "\", $STR_NOCASESENSEBASIC, -1))

        ; Check if the extension .au3 is appended to the end of the filename.
        Local $iExtension = StringInStr($sFileName, ".", $STR_NOCASESENSEBASIC)

        ; If a period (dot) is found then check whether or not the extension is equal to .au3.
        If $iExtension Then
            ; If the extension isn't equal to .au3 then append to the end of the filepath.
            If Not (StringTrimLeft($sFileName, $iExtension - 1) = ".csv") Then $sFileSaveDialog &= ".csv"
        Else
            ; If no period (dot) was found then append to the end of the file.
            $sFileSaveDialog &= ".csv"
        EndIf

        ; Display the saved file.
        MsgBox($MB_SYSTEMMODAL, "", "You saved the following file:" & @CRLF & $sFileSaveDialog)
    EndIf

	Local $sFileRead = _ArrayToString($aCSVRows,",")

	$hFileOpen = FileOpen($sFileSaveDialog, $FO_OVERWRITE)
    If $hFileOpen = -1 Then
        MsgBox($MB_SYSTEMMODAL, "", "An error occurred whilst writing the temporary file. :" &@error)
        Return False
    EndIf

    ; Write data to the file using the handle returned by FileOpen.
    FileWrite($hFileOpen, $sFileRead)
    ; Close the handle returned by FileOpen.
    FileClose($hFileOpen)


    EndIf
EndFunc   ;==>Main
