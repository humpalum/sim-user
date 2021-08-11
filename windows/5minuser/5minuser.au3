;#NoTrayIcon
#AutoIt3Wrapper_Change2CUI=y
#include <Array.au3>
#include <WinAPI.au3>
#include <Word.au3>
#include <IE.au3>

Opt("SendKeyDelay", 85)
; define global task list
Global $aTasks[3] = ['PowerShell_0', 'WordDocument_2', 'InternetExplorer_4'] 
;Global $aTasks[3] = ['PowerShell_0', 'InternetExplorer_4'] 
; creates Sleep Times array
Global $aSleepTimes[4] = ['39843', '60650', '121091', '78416']
;Global $aSleepTimes[4] = ['10000', '10000', '10000', '10000']
; copies original array just encase the task list borks
$aRandTasks = $aTasks

$aRandSleepTimes = $aSleepTimes
; shuffle this array so it's unique everytime
; it gets baked into the file
;_ArrayShuffle($aRandSleepTimes)

For $i In $aTasks
    ; start with a sleep Value
    ; pops the last shuffled value from the sleep array and assigns
    Local $vSleepTime =_ArrayPop($aRandSleepTimes)
    ConsoleWrite("[!] I will now sleep for : " & $vSleepTime & @CRLF)
    ; Sets the sleep time
    Sleep($vSleepTime)
    ;ConsoleWrite($i & @CRLF)
    ; gets the current function from the shuffled array
    $curfunc = ($i & @CRLF)
    ConsoleWrite($curfunc)
    ; call the function from the shuffled array
    ; the magic call
    Call($i)

Next

; < ----------------------------------- >
; <      PowerShell Interaction
; < ----------------------------------- >

PowerShell_0()

Func PowerShell_0()

    ; Creates a PowerShell Interaction

    Send("#r")
    ; Wait 10 seconds for the Run dialogue window to appear.
    WinWaitActive("Run", "", 10)
    ; note this needs to be escaped
    Send('powershell{ENTER}')
    ; check to see if we are already in an RDP session
    $active_window = _WinAPI_GetClassName(WinGetHandle("[ACTIVE]"))
    ConsoleWrite($active_window & @CRLF)
    $inRDP = StringInStr($active_window, "TscShellContainerClass")
    ; if the result is greater than 1 we are inside an RDP session
    if $inRDP < 1 Then
        WinWaitActive("Windows PowerShell", "", 10)
        SendKeepActive("Windows PowerShell")
    EndIf


    Send("gwmi win32_service{ENTER}")
    sleep(11112)
    Send("$psversiontable{ENTER}")
    sleep(12201)
    Send("ping 8.8.8.8{ENTER}")
    sleep(14820)
    Send("gci{ENTER}")
    sleep(6283)
    Send("ipconfig /all{ENTER}")
    sleep(4559)
    Send("netstat -anto{ENTER}")
    sleep(16720)
    Send("set-content test.txt LoremIpsumAndSoOn{ENTER}")
    sleep(16941)
    Send("get-content test.txt{ENTER}")
    sleep(9172)
    Send('exit{ENTER}')
    ; Reset Focus
    SendKeepActive("")

EndFunc


; < ----------------------------------- >
; <         Word Interaction
; < ----------------------------------- >


WordDocument_2()



Func WordDocument_2()
    ConsoleWrite("Check if Word is installed")
    Local $objWord = ObjCreate("Word.Application")

    If IsObj($objWord) then 
        ConsoleWrite("Version: " & $objWord.Caption & " " & $objWord.Version & @crlf & "Build: " & $objWord.Build)
    $objWord.Quit
    Else
        ConsoleWrite("Word is not installed\n")
        return
    EndIf
    ; Creates a Word Document : %USERPROFILE%\Invoice.docx

    Local $oWord = _Word_Create()

    ; Add a new empty document
    $oDoc = _Word_DocAdd($oWord)

    WinActivate("[CLASS:OpusApp]")
    WinWaitActive("[CLASS:OpusApp]")
    SendKeepActive("[CLASS:OpusApp]")


    Send("Lorem ipsum dolor sit amet, consetetur sadipscing elitr, sed diam nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam erat, sed diam voluptua. At vero eos et accusam et justo duo dolores et ea rebum. Stet clita kasd gubergren, no sea takimata sanctus est Lorem ipsum dolor sit amet. Lorem ipsum dolor sit amet, consetetur sadipscing elitr, sed diam nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam erat, sed diam voluptua. At vero eos et accusam et justo duo dolores et ea rebum. Stet clita kasd gubergren, no sea takimata sanctus est Lorem ipsum dolor sit amet.")

    ; Reset the SendKeep Active
    SendKeepActive("")
    ; now save
    _Word_DocSaveAs($oDoc,'%USERPROFILE%\Invoice.docx', $WdFormatDocumentDefault)
    _Word_DocClose($oDoc)



    Send("!{F4}")

EndFunc


; < ------------------------------------------ >
;         InternetExplorer Interaction        
; < ------------------------------------------ >

InternetExplorer_4()


Func InternetExplorer_4()

    ; Creates a InternetExplorer Interaction

    Local $oIE = _IECreate("https://www.youtube.com/watch?v=dQw4w9WgXcQ",1,1,1)
    Sleep(2000)
    ;WinWaitActive("Windows Internet Explorer")
    ;SendKeepActive("Windows Internet Explorer")
    ;WinSetState("Windows Internet Explorer","",@SW_MAXIMIZE)
    ; hardcoded sleep for now
    ; will convert to AutoIT random
    ; this is also where the IE interaction such as logging in etc will happen,
    ; spawning new tabs etc
    ; prob need a call out function to trigger a subroutine
    Sleep(20000)
    Send("!{F4}")



SendKeepActive("")
_IEQuit($oIE)

EndFunc

