#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <GuiListView.au3>
#include <Array.au3>
#include <File.au3>
#include <Word.au3>

Global $wdPropertyWords = 15

#Region ### START Koda GUI section ### Form=
$mainForm = GUICreate("Get Word Count For multiple word files - Free Edition - Advanced Office Automation", 1100, 600)
$selectButton = GUICtrlCreateButton("Select", 530, 8, 150, 25)
$fileListView = GUICtrlCreateListview("", 8, 50, 1080, 515, BitOR($WS_HSCROLL,$WS_VSCROLL))
_GUICtrlListView_InsertColumn($fileListView, 0, "Word Count", 182,2)
_GUICtrlListView_InsertColumn($fileListView, 0, "File", 900,2)
GUICtrlSetLimit ( $fileListView, 10000000)
$fileListContextMenu = GUICtrlCreateContextMenu ( $fileListView )

$folderPathInput = GUICtrlCreateInput("", 88, 10, 420, 20)
$pathLabel = GUICtrlCreateLabel("Folder Path", 10, 15)
$statusLabel = GUICtrlCreateLabel("......", 10, 570,100)
$includeSubfoldersCheckbox = GUICtrlCreateCheckbox("Include files in Subfolders", 700, 10, 185, 25)


GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###



While 1
    $guiMessage = GUIGetMsg()
    Switch $guiMessage
        Case $GUI_EVENT_CLOSE
            Exit
        Case $selectButton
            $selectedFolder = FileSelectFolder ("Choose word folder", @ScriptDir)
            GUICtrlSetData($folderPathInput, $selectedFolder)
            GetFilesList($selectedFolder, "*.docx;*.doc", isCheckboxChecked($includeSubfoldersCheckbox)) ;add extensions as needed separated by semicolon
    EndSwitch
WEnd

Func GetWordCount($file, $oWord)
    $oDoc = _Word_DocOpen($oWord, $file, Default, Default, True) ; Open in read-only mode
    If @error Then
        MsgBox(16, "Word UDF: _Word_DocOpen Example", "Error opening '" & $file & "'.")
        Return -1
    Else
        $wordCount = $oDoc.BuiltInDocumentProperties($wdPropertyWords).Value
        _Word_DocClose($oDoc) ; Close the document after getting the word count
        Return $wordCount
    EndIf
EndFunc
;~ ======================================== Get the files list based on word files extension *.docx ===========================================
Func GetFilesList ($Path = @ScriptDir, $Type = "*", $recurs = $FLTAR_NORECUR)
Local $recur_val
Local $FPath = $Path

    if $Type = "" Then $Type = "*" ; the type is specified in in the user interface, however it can be amended to accept other extensions like .doc - .docm
    if $recurs = True Then
        $recursionFlag = $FLTAR_RECUR
    Else
        $recursionFlag  = $FLTAR_NORECUR
    EndIf
    if $Path <> "" Then
            GUICtrlSetData($statusLabel,"Getting files....")
            GUICtrlSetState($selectButton, 128)
            Local $fileList = _FileListToArrayRec ($Path,$Type,1,$recursionFlag,Default,$FLTAR_FULLPATH)

        if UBound ($fileList) > 1 Then
            if @error Then MsgBox (0,"","Error code: " & @error & " - Extended code: " & @extended)

            deleteListViewItems($fileListView)
            Local $wordApp = _Word_Create(False)
            For $i = 1 To $fileList[0]
                If StringInStr($fileList[$i], "$") = 0 Then ; Check if the file name includes a dollar sign
                    Local $wordCount = GetWordCount($fileList[$i], $wordApp)
                    If $wordCount <> -1 Then
                        GUICtrlCreateListViewItem($fileList[$i] & "|" & $wordCount, $fileListView)
                    EndIf
                EndIf
            Next
             _Word_Quit($wordApp)
            GUICtrlSetData($statusLabel,"Done Getting files")

        Else
            GUICtrlSetData($statusLabel,"No files found")
            MsgBox (16,"Error","Cannot find any files")

        EndIf
        GUICtrlSetState($selectButton, 64)
    EndIf
EndFunc;==================End of the GetFilesList function ============>



;~ ======================== Check if the checkbox is checked or not and return true if it is checked ===========================================
Func isCheckboxChecked($checkboxControlID)
    Return BitAND(GUICtrlRead($checkboxControlID), $GUI_CHECKED) = $GUI_CHECKED
EndFunc ;=================== End of isCheckboxChecked ==============>



;~ =========================================== Delete item in the files list ===================
Func deleteListViewItems($listControlID)
    _GUICtrlListView_DeleteAllItems($listControlID)
EndFunc ;===================== End of deleteListViewItems function =========================>


