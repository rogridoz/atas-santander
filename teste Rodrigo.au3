#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----T004------------------------------------------------------------------------

; Script Start - Add your code below here
#include <Excel.au3>

Global $CNPJ[1400]
Global $DATANASC[1400]

GLOBAL $Hotkey
Global $n = 0
Global $i = 1
HotKeySet("{F3}","Pause")
HotKeySet("{F4}","Kill")
$sFilePath1 = "C:\Users\T780000\Desktop\VALIDACAOMEI.xlsx" ;This file should already exist
$oExcel = _ExcelBookOpen($sFilePath1)

#comments-start - DESCRIÇÃO DAS FUNÇÕES:

INICIO - BUSCAR CNPJ
PAUSE - PAUSAR ROBÔ
KILL - ENCERRAR O ROBÔ

#comments-end

For $i = 1 To 100;Loop
    $sCellValue = _ExcelReadCell($oExcel, $i, 1)
	$CNPJ[$n] = $sCellValue
	$cCellValue = _ExcelReadCell($oExcel, $i, 2)
	$DATANASC[$n] = $cCellValue
	
	$n = $n + 1
Next
	$n = 0
	Sleep(200)	

;SCRIPT VALIDACAO MEI
For $n = 0 To 100

MouseClick("LEFT", 99, 12) ;clica no botão do navegador
Sleep(200)
MouseClick("LEFT", 224, 56) 

;Colocar CNPJ
MouseClick("LEFT", 219, 338)
For $i = 1 To 15
Send("{DEL}")
Next
Send($CNPJ[$n])

Sleep(1200)
MouseClick("LEFT", 677, 338)
Sleep(1200)
For $i = 1 To 15
Send("{DEL}")
Next
Send($DATANASC[$n])

Sleep(8200)



MouseClick("LEFT", 677, 338)
Sleep(8200)


next

FUNC PAUSE()
	$HotKey = NOT $Hotkey
	
	While $HotKey
			sleep(500)
			ToolTip('Robô está "Pausado". ||| Para finalizar o Robô : F4' ,0,0)
		WEnd
		ToolTip("")
	EndFunc
	
Func Kill()
	msgBox(0,"Fim","O robô foi encerrado com sucesso.",1)
	
	
	Exit 0
EndFunc
