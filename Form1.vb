Option Explicit On
Imports System.Drawing.Printing
Imports System.IO
Imports System.IO.File
Imports System.Runtime.Serialization

' Error numbering:
' 1xxx  Form1
' 2xxx  DispRem
' 3xxx  Find
' 4xxx  FindAll
' 5xxx  GridSearch
' 6xxx  OneSymp (Repertory-style search)
' 7xxx  PrescRem
' 8xxx  Questionnaire
' 9xxx  SelRem
' 10xxx SetAutowt
' 11xxx SetPref

Public Class Form1
    Public Const InFile = 1   ' symptextdata.dat; 1st line is symptom description; alternate lines are data which contain symptom ID, Remedy ID, numbered occurrence of this remedy in the symptom list (symptom ID within the remedy), [next Remedy ID, symptom ID within the remedy] etc.
    Public Const CaseFile = 2  ' user's case file
    Public Const SdxFile = 7  ' seartxtidx.dat (words + pointers interleaved)
    Public Const QuestFile = 8  ' user's questionaire file
    Public Const PrescFile = 9    ' Prescribe export file
    Public Const NumFiles = 9      ' Number of different files

    Public Const CombSympPtr = 1344716 ' byte pointer for InFile
    Public Const SCombSympPtr = 76893 ' byte pointer for SdxFile
    Public Const file_size = 25395   ' # of lines in symptoms file
    Public Const DHA_Version = "30"    ' software version number

    Public Shared Afile_size As Long   ' # of lines in AutoWt.dat file
    Public Shared AlphaCombSympPtr As String   ' string version of CombSympPtr
    Public Shared AutoWeight As Boolean    ' automatic symptom weight option flag
    Public Shared AutoWeightTp As Integer ' type from auto-weight type pop-up 1 = physical, 2 = general, 3 = mind
    Public Shared AutoWeightSev As Integer ' severity from auto-weight severity pop-up 1 = low, 2 = medium, 3 = high
    Public Shared AutoWeightText As String ' text for auto-weight pop-up
    Public Shared CalledFromShowSymSymps As Boolean ' True if called from ShowSymilarSymptoms, False otherwise
    Public Shared CaseFileName As String
    Public Shared CaseFilePath As String
    Public Shared CasePrint As Boolean   ' case print flag from PrintSelections Dialog
    Public Shared CombSympData As New ArrayList()  ' Holds ListBox symptom strings and ID data
    ' Symptom String:
    ' 1st number Is symptom ID
    ' 2nd number Is Remedy ID
    ' 3rd number Is numbered occurrence of this remedy in the combsymp list (symptom ID within the remedy)
    Public Shared Dirty As Boolean       ' data has changed since saving file flag
    Public Shared DisplayCount As Integer ' # of remedies to display
    Public Shared ErrorFlag As Integer     ' error message number
    Public Shared ExitPressed As Boolean   ' intro form EXIT button was pressed, bag out!
    Public Shared file_pos As Long    ' position of Lst1 in symptom file
    Public Shared FileExitCalled As Boolean ' Denotes that FileExit function has been called
    Public Shared FindAllCaller As Integer ' 1 = Find, 2 = Grid, 3 = Quest, 4 = Show Similar Symptoms
    Public Shared FindAllVisible As Boolean  ' form FindAll is visible
    Public Shared FindLoaded As Boolean     ' Used to keep Find_Load from being executed more than once
    Public Shared FirstLoad As Boolean     ' Used to suppress adding first list box item (that gets selected when loading) from being added to SelectList
    Public Shared FirstRepSearch As Boolean ' Used to suppress re-init of Load items
    Public Shared FirstQuestFileSave As Boolean    ' need to confirm name flag
    Public Shared GlobalSearchData() As String  ' Holds data from file seartxtidx.dat
    Public Shared GridString(24) As String  ' used by grid search routine
    Public Shared InLoad As Boolean    ' Used to keep from resetting First until after Form1_Load exits
    Public Shared MapSymIDToIndex() As Integer  ' array elements correspond to symptom ID's, and contain Lst1 index
    Public Shared MustBoxSelected As Boolean    ' keeps track of selection status of text in MustBox
    Public Shared MustList(1) As Integer ' list of "must use" selected symptom ID's
    Public Shared MustData As String   ' CombSympData.SympData string associated with MustBox SympDesc (Text)
    Public Shared MustListSize As Integer  ' # of items in "must use" select list
    Public Shared Normalize As Boolean ' normalize prescribe numbers flag
    Public Shared NoSelectedItems As Boolean    ' Used to cause (re)-setting of SelLst data source when first item is added
    Public Shared pageNo As Integer        ' Keep track of page printing in case printing only a selection of pages
    Public Shared Pres(500) As String     'Prescribe string
    Public Shared PrintPreviewShow As Boolean  ' Tells FilePrint() to show printPreviewDialog instead of printDialog
    Public Shared QDirty As Boolean        ' questionaire contains unsaved data flag
    Public Shared QuestFileName As String
    Public Shared QuestFilePath As String
    Public Shared QuestFirstLoad As Boolean    ' Questionnaire first time load; prevents deleting data if re-loading
    Public Shared QuestionnaireVisible As Boolean   ' Questionnaire is visible (has been loaded) flag
    Public Shared QPrint As Boolean ' Questionnaire print flag from PrintSelections Dialog
    Public Shared QuestWeight As Integer           ' weight from questionaire
    Public Shared RemedyFileName As String    ' name of remedies file
    Public Shared RemedyNum As Integer    ' line # from remedy list
    Public Shared RemPrint As Boolean    ' remedy print flag, from PrintSelections Dialog
    Public Shared RemText() As String    ' array list to contain RemPresc.dat file Text
    Public Shared RemData() As String    ' array list to contain RemPresc.dat file Data
    Public Shared RepertorySearch As Boolean   ' leave repertory popped-up
    Public Shared RaiseException As Boolean ' tells function caller to raise an exception
    Public Shared Remfile_size As Long   ' # of lines in RemPresc.dat file
    Public Shared SearchString As String  ' used by search routine
    Public Shared SelectionMode As Integer    ' 1 = default; 2 = equal percent; 3 = straight count
    Public Shared SelSympData As New ArrayList()  ' Holds SelLst symptom strings and ID data
    Public Shared SelRemData As New ArrayList()    ' Holds PrescRem ListBox1 remedy strings and ID data
    Public Shared Sfile_size As Long   ' # of lines in file seartxtidx.dat
    Public Shared ShowDisclaimer As Boolean ' shows disclaimer
    Public Shared ShowIntroduction As Boolean   ' shows introduction
    Public Shared ShowStatusbar As Boolean ' shows status bar
    Public Shared ShowToolbar As Boolean   ' shows toolbar
    Public Shared SMaxSympPtr As Long      ' maximum value of SCombSympPtr
    Public Shared stringBuffer As String   ' string buffer used by print event handler
    Public Shared Syfile_size As Long   ' # of lines in Thestxtidx.dat file
    Public Shared SympPrint As Boolean  ' symptom print flag from PrintSelections Dialog
    Public Shared SymSelected(1000) As Integer ' list of selected symptom ID's in Lst1
    Public Shared SymSelectedSize As Integer  ' # of selected symptom ID's in Lst1
    Public Shared TempRemedyNumb As String    ' current line remedy number
    Public Shared ToolText As Boolean      ' enable tool text

    Public Tab38 As String = StrDup(38, Chr(9)) ' Used to insert 38 tabs to "hide" item off to the right of list

    ' Preference settings from registration database.
    Public Shared ScreenBackColor As Color            ' list background color
    Public Shared ScreenForeColor As Color          ' list foreground color
    Public Shared PrinterFontName As Font     ' Printer font name
    Public Shared ScreenFontName As Font     ' Text font name
    Public Shared PrinterFontSize As String     ' Printer font size
    Public Shared ScreenFontSize As String     ' Text font size

    ' Font attributes strings
    Public Shared FontName As String
    Public Shared FontSize As String

    ' Index arrays
    Shared numIndexes As Double
    Shared ptr As Integer

    Public Shared SymptomData() = System.IO.File.ReadAllLines(AppContext.BaseDirectory + DATA_DIR + "symptextdata.dat")   'This needs to be here because SymptomData() is also used by FindAll and QuestProgress

    ' For SDate the array element is the symptom number.  For all others the first is element number in form
    ' and the second is symptom number; it needs to be in this order to use the "Redim Preserve" when adding
    ' symptoms.
    Structure SympData
        Public Shared SDate() As String
        Public Shared Title() As String
        Public Shared P1(,) As String
        Public Shared P1Cb(,) As Boolean
        Public Shared P2(,) As String
        Public Shared P3(,) As String
        Public Shared P3Cb(,) As Boolean
        Public Shared P4Cb(,) As Boolean
        Public Shared P4(,) As String
    End Structure

    Public Shared Symptom As SympData
    Public Shared SymPtr As Integer   ' pointer to Symptom array element
    Public Shared SympSize As Integer  ' size of Symptom array

    Structure PatientData
        Public Shared Name As String
        Public Shared PDate As String
        Public Shared DOB As String
        Public Shared P5AgeCb() As Boolean
        Public Shared P5SexCb() As Boolean
        Public Shared P5() As String
        Public Shared P6() As String
        Public Shared P6Cb() As Boolean
        Public Shared P7() As String
        Public Shared P7Cb() As Boolean
    End Structure

    Public Shared Record As PatientData

    ' Variables "global" to Form1 only
    Dim AltPressed As Boolean   ' <ALT> key is currently pressed
    Dim CtrlPressed As Boolean  ' <CTRL> key is currently pressed
    Dim MaxRemHits As Integer   ' maximum # of symptoms for same remedy
    Dim MinNumb As Integer  ' minimum number hits to consider a remedy
    Dim NumRem As Integer   ' # of remedies in RemList
    Dim LastSelectedDataString As String    ' Lst1 data from latest selected item
    Dim RemList(1000) As Integer    ' list of remedy ID's
    Dim selectedFile As String  ' current working file

    Public Const CTRL_KEY As Integer = 17
    Public Const ALT_KEY As Integer = 18
    Public Const B_KEY As Integer = 66  ' Browse Remedies
    Public Const F_KEY As Integer = 70  ' Find
    Public Const G_KEY As Integer = 71  ' Grid search
    Public Const I_KEY As Integer = 73  ' SImilar symptoms
    Public Const M_KEY As Integer = 77  ' Must Use
    Public Const Q_KEY As Integer = 81  ' Questionnaire
    Public Const R_KEY As Integer = 82  ' Repertory-style search
    Public Const S_KEY As Integer = 83  ' Select Symptoms
    Public Const DATA_DIR As String = ".\Data\"
    Public Const HELP_DIR As String = ".\Help\"
    Public Const MAN_DIR As String = ".\Manual\"
    Public Property Printer As Object

    Public Function SymptomIsSelected(SympData As String)  ' Test to see if symptom is selected.
        Dim alreadySelected As Boolean  ' symptom is already selected flag
        Dim done As Boolean     ' found symptom, force loop exit
        Dim intID As Integer    ' integer form of Symptom ID
        Dim selPtr As Integer   ' pointer to selected symptom array

        alreadySelected = False
        ' Check that symptom not already selected.
        alreadySelected = False
        intID = Val(SympData)    ' Convert it to integer.
        done = False
        For selPtr = 0 To SelLst.Items.Count - 1
            If intID = Val(SelLst.Items(selPtr).SelSympData.ToString) Then
                alreadySelected = True
                done = True
            End If
            If done Then Exit For
        Next selPtr
        SymptomIsSelected = alreadySelected
    End Function


    Public Sub ClearQuestionaire()
        ' Clear questionaire object and any visible forms.

        ' Struct element arrays cannot have size at declaration time, so it needs to be done here; this gets executed at startup.
        ReDim Symptom.SDate(SympSize - 1)  ' Note that ReDim will assure that items are cleared.
        ReDim Symptom.Title(SympSize - 1)
        ReDim Symptom.P1(6, SympSize - 1)
        ReDim Symptom.P1Cb(4, SympSize - 1)
        ReDim Symptom.P2(6, SympSize - 1)
        ReDim Symptom.P3(1, SympSize - 1)
        ReDim Symptom.P3Cb(49, SympSize - 1)
        ReDim Symptom.P4Cb(4, SympSize - 1)
        ReDim Symptom.P4(1, SympSize - 1)
        ReDim Record.P5AgeCb(4)
        ReDim Record.P5SexCb(1)
        ReDim Record.P5(5)
        ReDim Record.P6(5)
        ReDim Record.P6Cb(14)
        ReDim Record.P7(14)
        ReDim Record.P7Cb(45)
        SympSize = 1
        SymPtr = 0

        ' Now clear questionaires data
        If QuestionnaireVisible Then
            WriteQ1Screen()
            WriteQ2Screen()
            WriteQ3Screen()
            WriteQ4Screen()
            WriteQ5Screen()
            WriteQ6Screen()
            WriteQ7Screen()
        End If
    End Sub

    Public Sub CloseAll()
        Dim FileNo As Integer

        Try
            For FileNo = 1 To NumFiles
                FileSystem.FileClose(FileNo)
            Next
        Catch
        End Try

        Try
            If FileLen(QuestFileName) = 0 Then Kill(QuestFilePath + QuestFileName)
        Catch
        End Try
        Try
            If FileLen(CaseFileName) = 0 Then Kill(CaseFileName)
        Catch
        End Try

    End Sub


    Public Function OpenFileWrite(FileName As String, FileNumber As Integer)
        Dim FileStatus As Boolean

        FileStatus = True   ' Assume success.
        Err.Description = ""
        FileSystem.FileOpen(FileNumber, FileName, OpenMode.Output)
        If Err.Description = "File already open" Then
            FileSystem.FileClose(FileNumber)
            Err.Description = ""
            FileSystem.FileOpen(FileName, FileNumber, OpenMode.Output)
            If Err.Description = "File Already Open" Then FileStatus = False
        End If
        Return (FileStatus)
    End Function

    Public Function OpenFileReadBinary(FileName As String, FileNumber As Integer)
        Dim FileStatus As Boolean

        FileStatus = True   ' Assume success.
        Err.Description = ""
        FileSystem.FileOpen(FileName, FileNumber, OpenMode.Binary, OpenAccess.Read)
        If Err.Description = "File already open" Then
            FileSystem.FileClose(FileNumber)
            Err.Description = ""
            FileSystem.FileOpen(FileName, FileNumber, OpenMode.Input)
            If Err.Description = "File Already Open" Then FileStatus = False
        End If
        Return (FileStatus)
    End Function
    Public Function OpenFileRead(FileName As String, FileNumber As Integer)
        Dim FileStatus As Boolean

        FileStatus = True   ' Assume success.
        Err.Description = ""
        FileSystem.FileOpen(FileNumber, FileName, OpenMode.Input)
        If Err.Description = "File already open" Then
            FileSystem.FileClose(FileNumber)
            Err.Description = ""
            FileSystem.FileOpen(FileName, FileNumber, OpenMode.Input)
            If Err.Description = "File Already Open" Then FileStatus = False
        End If
        Return (FileStatus)
    End Function

    Public Sub SetupAndPrint()
        ' Get printer font from registry
        PrinterFontName = New Font(GetSetting("Hahnasst", "Startup", "PrinterFontName", [Default]:="Microsoft Sans Serif"),
            GetSetting("Hahnasst", "Startup", "PrinterFontSize", [Default]:="8"))

        ' Apply settings from PrintDialog1
        PrintDocument1.PrinterSettings.PrinterName = PrintDialog1.PrinterSettings.PrinterName
        PrintDocument1.PrinterSettings.Copies = PrintDialog1.PrinterSettings.Copies
        PrintDocument1.PrinterSettings.PrintToFile = PrintDialog1.PrinterSettings.PrintToFile
        pageNo = 0
        PrintDocument1.PrinterSettings.ToPage = PrintDialog1.PrinterSettings.ToPage
        PrintDocument1.PrinterSettings.FromPage = PrintDialog1.PrinterSettings.FromPage
        PrintDialog1.Document = PrintDocument1

    End Sub

    Public Sub QuestPrint()
        ' Load the print buffer

        Dim LabelNo As Integer      ' label array pointer
        Dim msg As String           ' Error message string
        Dim S As String             ' line from list box

        Try
            ErrorFlag = 1001
            Cursor = Cursors.WaitCursor    ' Change cursor to hourglass.

            '   Print the page title.
            stringBuffer = "Patient Questionaire" + vbCrLf + vbCrLf

            ' Print tab 1
            stringBuffer += "Name:  " + PatientData.Name + "  Date:  " + PatientData.PDate + "  D.O.B.:  " + PatientData.DOB + vbCrLf
            For SymPtr = 0 To SympSize - 1
                stringBuffer += "Symptom #" + Str(SymPtr + 1)
                stringBuffer += Chr(9) + SympData.Title(SymPtr) + vbCrLf
                stringBuffer += Questionnaire.Label82.Text
                stringBuffer += Chr(9) + SympData.SDate(SymPtr) + vbCrLf
                stringBuffer += Questionnaire.Label5.Text
                stringBuffer += Chr(9) + SympData.P1(0, SymPtr) + vbCrLf
                stringBuffer += Questionnaire.Label6.Text
                stringBuffer += Chr(9) + SympData.P1(1, SymPtr) + vbCrLf
                stringBuffer += Questionnaire.Label7.Text
                stringBuffer += Chr(9) + SympData.P1(2, SymPtr) + vbCrLf
                stringBuffer += Questionnaire.Label8.Text
                stringBuffer += Chr(9) + SympData.P1(3, SymPtr) + vbCrLf
                stringBuffer += Questionnaire.Label9.Text
                stringBuffer += Chr(9) + SympData.P1(4, SymPtr) + vbCrLf
                stringBuffer += Questionnaire.Label10.Text
                stringBuffer += Chr(9) + SympData.P1(5, SymPtr) + vbCrLf
                stringBuffer += Questionnaire.Label11.Text
                If SympData.P1Cb(0, SymPtr) Then
                    S = Chr(9) + "yes"
                Else
                    S = Chr(9) + "no"
                End If
                stringBuffer += S + vbCrLf
                stringBuffer += Questionnaire.Label12.Text
                stringBuffer += Chr(9) + SympData.P1(6, SymPtr) + vbCrLf
                stringBuffer += Questionnaire.Label13.Text
                If SympData.P1Cb(4, SymPtr) Then
                    S = Chr(9) + "high"
                ElseIf SympData.P1Cb(3, SymPtr) Then
                    S = Chr(9) + "medium"
                Else
                    S = Chr(9) + "low"
                End If
                stringBuffer += S + vbCrLf

                ' Print tab 2
                stringBuffer += Questionnaire.Label14.Text
                stringBuffer += Chr(9) + SympData.P2(0, SymPtr) + vbCrLf
                stringBuffer += Questionnaire.Label15.Text
                stringBuffer += Chr(9) + SympData.P2(1, SymPtr) + vbCrLf
                stringBuffer += Questionnaire.Label16.Text
                stringBuffer += Chr(9) + SympData.P2(2, SymPtr) + vbCrLf
                stringBuffer += Questionnaire.Label17.Text
                stringBuffer += Chr(9) + SympData.P2(3, SymPtr) + vbCrLf
                stringBuffer += Questionnaire.Label18.Text
                stringBuffer += Chr(9) + SympData.P2(4, SymPtr) + vbCrLf
                stringBuffer += Questionnaire.Label19.Text
                stringBuffer += Chr(9) + SympData.P2(5, SymPtr) + vbCrLf
                stringBuffer += Questionnaire.Label20.Text
                stringBuffer += Chr(9) + SympData.P2(6, SymPtr) + vbCrLf

                ' Print tab 3
                stringBuffer += Questionnaire.Label21.Text
                stringBuffer += Chr(9) + SympData.P3(0, SymPtr) + vbCrLf
                stringBuffer += Questionnaire.Label22.Text
                stringBuffer += Chr(9) + SympData.P3(1, SymPtr) + vbCrLf
                stringBuffer += Questionnaire.Label23.Text
                If SympData.P3Cb(0, SymPtr) Then
                    S = Chr(9) + "winter"
                ElseIf SympData.P3Cb(1, SymPtr) Then
                    S = Chr(9) + "spring"
                ElseIf SympData.P3Cb(2, SymPtr) Then
                    S = Chr(9) + "summer"
                Else
                    S = Chr(9) + "fall"
                End If
                stringBuffer += S + vbCrLf
                stringBuffer += Questionnaire.Label24.Text
                If SympData.P3Cb(4, SymPtr) Then
                    S = Chr(9) + "winter"
                ElseIf SympData.P3Cb(5, SymPtr) Then
                    S = Chr(9) + "spring"
                ElseIf SympData.P3Cb(6, SymPtr) Then
                    S = Chr(9) + "summer"
                Else
                    S = Chr(9) + "fall"
                End If
                stringBuffer += S + vbCrLf
                stringBuffer += Questionnaire.Label25.Text + vbCrLf

                For LabelNo = 0 To 9
                    S = Chr(9) + Questionnaire.Label(LabelNo).Text
                    If SympData.P3Cb((LabelNo * 3) + 8, SymPtr) Then
                        S = S + " improves"
                    ElseIf SympData.P3Cb((LabelNo * 3) + 9, SymPtr) Then
                        S = S + " worsens"
                    Else
                        S = S + " no change"
                    End If
                    stringBuffer += S + vbCrLf
                Next LabelNo
                stringBuffer += Questionnaire.Label36.Text + vbCrLf
                For LabelNo = 10 To 13
                    S = Chr(9) + Questionnaire.Label(LabelNo + 1).Text
                    If SympData.P3Cb((LabelNo * 3) + 8, SymPtr) Then
                        S = S + " improves"
                    ElseIf SympData.P3Cb((LabelNo * 3) + 9, SymPtr) Then
                        S = S + " worsens"
                    Else
                        S = S + " no change"
                    End If
                    stringBuffer += S + vbCrLf
                Next LabelNo

                ' Print tab 4
                stringBuffer += Questionnaire.Label41.Text
                If SympData.P4Cb(0, SymPtr) Then
                    S = Chr(9) + "standing"
                ElseIf SympData.P4Cb(1, SymPtr) Then
                    S = Chr(9) + "sitting"
                ElseIf SympData.P4Cb(2, SymPtr) Then
                    S = Chr(9) + "lying"
                Else
                    S = Chr(9) + "no difference"
                End If
                stringBuffer += S + vbCrLf
                stringBuffer += Questionnaire.Label42.Text
                stringBuffer += Chr(9) + SympData.P4(0, SymPtr) + vbCrLf
                stringBuffer += Questionnaire.Label43.Text
                stringBuffer += Chr(9) + SympData.P4(1, SymPtr) + vbCrLf + vbLf
            Next SymPtr

            ' Print tab 5
            stringBuffer += "Personal Characteristics" + vbCrLf
            stringBuffer += Questionnaire.Label44.Text
            If PatientData.P5AgeCb(0) Then
                S = Chr(9) + "infant"
            ElseIf PatientData.P5AgeCb(1) Then
                S = Chr(9) + "child"
            ElseIf PatientData.P5AgeCb(2) Then
                S = Chr(9) + "adolescent"
            ElseIf PatientData.P5AgeCb(3) Then
                S = Chr(9) + "adult"
            Else
                S = Chr(9) + "senior"
            End If
            stringBuffer += S
            If PatientData.P5SexCb(0) Then
                S = Chr(9) + "male"
            Else
                S = Chr(9) + "female"
            End If
            stringBuffer += S + vbCrLf
            stringBuffer += Questionnaire.Label45.Text
            stringBuffer += Chr(9) + PatientData.P5(0) + vbCrLf
            stringBuffer += Questionnaire.Label46.Text
            stringBuffer += Chr(9) + Chr(9) + PatientData.P5(1) + vbCrLf
            stringBuffer += Questionnaire.Label47.Text
            stringBuffer += Chr(9) + PatientData.P5(2) + vbCrLf
            stringBuffer += Questionnaire.Label48.Text
            stringBuffer += Chr(9) + PatientData.P5(3) + vbCrLf
            stringBuffer += Questionnaire.Label49.Text
            stringBuffer += Chr(9) + PatientData.P5(4) + vbCrLf
            stringBuffer += Questionnaire.Label50.Text
            stringBuffer += Chr(9) + PatientData.P5(5) + vbCrLf

            ' Print tab 6
            stringBuffer += Questionnaire.Label51.Text
            stringBuffer += Chr(9) + PatientData.P6(0) + vbCrLf
            stringBuffer += Questionnaire.Label52.Text
            stringBuffer += Chr(9) + PatientData.P6(1) + vbCrLf
            stringBuffer += Questionnaire.Label53.Text
            If PatientData.P6Cb(0) Then
                S = S + " low"
            ElseIf PatientData.P6Cb(1) Then
                S = S + " medium"
            Else
                S = S + " high"
            End If
            stringBuffer += S + vbCrLf
            stringBuffer += Questionnaire.Label57.Text
            stringBuffer += Chr(9) + PatientData.P6(2)
            stringBuffer += Questionnaire.Label53.Text
            If PatientData.P6Cb(3) Then
                S = S + " low"
            ElseIf PatientData.P6Cb(4) Then
                S = S + " medium"
            Else
                S = S + " high"
            End If
            stringBuffer += S + vbCrLf
            stringBuffer += Questionnaire.Label58.Text
            stringBuffer += Chr(9) + PatientData.P6(3)
            stringBuffer += Questionnaire.Label53.Text
            If PatientData.P6Cb(6) Then
                S = S + " low"
            ElseIf PatientData.P6Cb(7) Then
                S = S + " medium"
            Else
                S = S + " high"
            End If
            stringBuffer += S + vbCrLf
            stringBuffer += Questionnaire.Label59.Text
            stringBuffer += Chr(9) + PatientData.P6(4)
            stringBuffer += Questionnaire.Label53.Text
            If PatientData.P6Cb(9) Then
                S = S + " low"
            ElseIf PatientData.P6Cb(10) Then
                S = S + " medium"
            Else
                S = S + " high"
            End If
            stringBuffer += S + vbCrLf
            stringBuffer += Questionnaire.Label60.Text
            stringBuffer += Chr(9) + PatientData.P6(5)
            stringBuffer += Questionnaire.Label53.Text
            If PatientData.P6Cb(12) Then
                S = S + " low"
            ElseIf PatientData.P6Cb(13) Then
                S = S + " medium"
            Else
                S = S + " high"
            End If
            stringBuffer += S + vbCrLf

            ' Print tab 7
            For LabelNo = 0 To 14
                S = Questionnaire.Label(LabelNo + 14).Text + Chr(9) + PatientData.P7(LabelNo) + Chr(9) + "severity "
                If PatientData.P7Cb(3 * LabelNo) Then
                    S = S + " low"
                ElseIf PatientData.P7Cb((3 * LabelNo) + 1) Then
                    S = S + " medium"
                Else
                    S = S + " high"
                End If
                stringBuffer += S + vbCrLf
            Next LabelNo

            Cursor = Cursors.Default    ' Change cursor to default.

        Catch
            Cursor = Cursors.Default    ' Change cursor to default.
            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub

    ' Print event handler
    Sub printDocument1_PrintPage(ByVal sender As Object,
    ByVal e As PrintPageEventArgs) Handles PrintDocument1.PrintPage

        Dim charactersOnPage As Integer = 0
        Dim linesPerPage As Integer = 0
        Dim myStringFormat As StringFormat = New StringFormat()
        Dim tabStops() As Single = {25.0, 25.0, 25.0}   ' Define 3 tab stops
        Dim locMarginBounds As Rectangle    ' Used to modify page margins

        myStringFormat.SetTabStops(0.0, tabStops)
        pageNo += 1

        ' Sets the value of charactersOnPage to the number of characters 
        ' of stringBuffer that will fit within the bounds of the page.
        e.Graphics.MeasureString(stringBuffer, PrinterFontName, e.MarginBounds.Size,
        StringFormat.GenericTypographic, charactersOnPage, linesPerPage)
        e.PageSettings.PrinterSettings.FromPage = PrintDocument1.PrinterSettings.FromPage
        e.PageSettings.PrinterSettings.ToPage = PrintDocument1.PrinterSettings.ToPage
        ' If printing a range of pages, read down to the first page.
        While pageNo < PrintDocument1.PrinterSettings.FromPage And stringBuffer.Length > 0
            e.Graphics.MeasureString(stringBuffer, PrinterFontName, e.MarginBounds.Size,
        StringFormat.GenericTypographic, charactersOnPage, linesPerPage)
            stringBuffer = stringBuffer.Substring(charactersOnPage)
            pageNo += 1
        End While

        If stringBuffer.Length = 0 Then
            MsgBox("Nothing To print!", vbOKOnly)
            e.HasMorePages = False
            stringBuffer = stringBuffer
            Exit Sub
        End If

        ' For some reason, the MarginBounds sets the page height such that the upper half of
        ' the first line for the next page is printed after the last line on the current page.
        ' Need to introduce a fudge factor so this won't happen.
        locMarginBounds = e.MarginBounds
        locMarginBounds.Height -= (e.MarginBounds.Height / (linesPerPage + 0.5)) / 2

        ' Draws the string within the bounds of the page.
        e.Graphics.DrawString(stringBuffer, PrinterFontName, Brushes.Black,
                        locMarginBounds, myStringFormat)

        ' Remove the portion of the string that has been printed.
        stringBuffer = stringBuffer.Substring(charactersOnPage)

        ' Check to see if more pages are to be printed.
        e.HasMorePages = stringBuffer.Length > 0
        If pageNo = PrintDocument1.PrinterSettings.ToPage Then e.HasMorePages = False

        ' If there are no more pages, reset the string to be printed.
        If Not e.HasMorePages Then
            stringBuffer = ""
        End If

    End Sub

    Public Sub DisableFileClose()
        Me.CloseCaseFile.Enabled = False
        ToolStripFileClose.Enabled = False
    End Sub
    Public Sub DisableFileNew()
        Me.NewCaseFile.Enabled = False
        ToolStripFileNew.Enabled = False
    End Sub

    Public Sub DisableFileOpen()
        Me.OpenCaseFile.Enabled = False
        ToolStripFileOpen.Enabled = False
    End Sub

    Public Sub DisableFilePrintPreview()
        Me.PrintPreview.Enabled = False
    End Sub
    Public Sub DisableFilePrint()
        Me.Print.Enabled = False
        Me.PrintPreview.Enabled = False
        ToolStripPrint.Enabled = False
    End Sub

    Public Sub DisableFileQuestClose()
        Me.CloseQuestFile.Enabled = False
        Questionnaire.mnuCloseQuestFile.Enabled = False
    End Sub

    Public Sub DisableFileQuestNew()
        Me.NewQuestFile.Enabled = False
        Questionnaire.mnuNewQuestFile.Enabled = False
    End Sub

    Public Sub DisableFileQuestOpen()
        Me.OpenQuestFile.Enabled = False
        Questionnaire.mnuOpenQuestFile.Enabled = False
    End Sub

    Public Sub DisableFileQuestSave()
        Me.SaveQuestFile.Enabled = False
        Questionnaire.mnuSaveQuestFile.Enabled = False
    End Sub

    Public Sub DisableFileQuestSaveAs()
        Me.SaveQuestFileAs.Enabled = False
        Questionnaire.mnuSaveQuestFileAs.Enabled = False
    End Sub

    Public Sub DisableFileSave()
        Me.SaveCaseFile.Enabled = False
        ToolStripFileSave.Enabled = False
    End Sub

    Public Sub DisableFileSaveAs()
        Me.SaveCaseFileAs.Enabled = False
        ToolStripFileSaveAs.Enabled = False
    End Sub

    Public Sub DisableQuestFileClose()
        Me.CloseQuestFile.Enabled = False
        Questionnaire.mnuCloseQuestFile.Enabled = False
    End Sub

    Public Sub EnableFileClose()
        Me.CloseCaseFile.Enabled = True
        ToolStripFileClose.Enabled = True
    End Sub


    Public Sub EnableFileNew()
        Me.NewCaseFile.Enabled = True
        ToolStripFileNew.Enabled = True
    End Sub

    Public Sub EnableFileOpen()
        Me.OpenCaseFile.Enabled = True
        ToolStripFileOpen.Enabled = True
    End Sub

    Public Sub EnableFilePrintPreview()
        Me.PrintPreview.Enabled = True
    End Sub
    Public Sub EnableFilePrint()
        Me.Print.Enabled = True
        Me.PrintPreview.Enabled = True
        ToolStripPrint.Enabled = True
    End Sub

    Public Sub EnableFileQuestClose()
        Me.CloseQuestFile.Enabled = True
        Questionnaire.mnuCloseQuestFile.Enabled = True
    End Sub

    Public Sub EnableFileQuestNew()
        Me.NewQuestFile.Enabled = True
        Questionnaire.mnuNewQuestFile.Enabled = True
    End Sub

    Public Sub EnableFileQuestOpen()
        Me.OpenQuestFile.Enabled = True
        Questionnaire.mnuOpenQuestFile.Enabled = True
    End Sub

    Public Sub EnableFileQuestSave()
        Me.SaveQuestFile.Enabled = True
        Questionnaire.mnuSaveQuestFile.Enabled = True
    End Sub

    Public Sub EnableFileQuestSaveAs()
        Me.SaveQuestFileAs.Enabled = True
        Questionnaire.mnuSaveQuestFileAs.Enabled = True
    End Sub

    Public Sub EnableFileSave()
        SaveCaseFile.Enabled = True
        ToolStripFileSave.Enabled = True
    End Sub
    Public Sub EnableFileSaveAs()
        SaveCaseFileAs.Enabled = True
        ToolStripFileSaveAs.Enabled = True
    End Sub

    Public Sub FileClose()
        Dim msg As String   ' dialog box string

        Try
            ErrorFlag = 1002
            If (Dirty) Then ' Ask if user wants to save the file.
                msg = "Data has changed; do you want to save the current case file?"
                ' If data changed, allow user to save current file.
                If (MsgBox(msg, vbYesNo + vbQuestion) = 6) Then
                    If CaseFileName = "Untitled" Then
                        Call FileSaveAs()
                    Else
                        Call SaveCurrentFile()
                    End If
                End If
            End If

            FileSystem.FileClose(CaseFile)

            MustBox.Text = ""
            MustData = ""
            If NoSelectedItems = False Then ' This would cause an exception if data source was not already connected
                SelSympData.Clear()
                SelLst.DataSource = Nothing
                SelLst.DataSource = SelSympData
                SelLst.ValueMember = "SelSympData"
                SelLst.DisplayMember = "SelSympDesc"
                SelLst.ClearSelected()
                NoSelectedItems = True
            End If

            Lst1.SelectedIndices.Clear()
            SymSelectedSize = 0

            ErrorFlag = 1003
            SelectRemedy.Enabled = True
            EnableFileNew()
            EnableFileOpen()
            DisableFileSave()
            DisableFileSaveAs()
            DisableFileClose()
            DisableFilePrintPreview()
            DisableFilePrint()
            PrescribeButton.Enabled = False

            CaseFileName = "Untitled"
            WriteTitles()     ' Display file names at top of page.

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")

        End Try

    End Sub

    Public Sub FileExit()

        If Not FileExitCalled Then
            FileExitCalled = True
            ' Close any opened files
            FileClose()
            FileQuestClose()
            CloseAll()
            ' Close forms that were opened at Form1 Load (including Form1):
            OneSymp.Close()
            Find.Close()
            GridSearch.Close()
            SelRem.Close()
            Me.Close()
        End If
    End Sub


    Public Sub FileNew()
        Dim msg As String
        Dim ptr As Integer  ' used to initialize symSelected array

        Try
            ErrorFlag = 1004
            If (Dirty) Then  ' Ask if user wants to save the file.
                msg = "Data has changed; do you want to save the current case file?"
                If (MsgBox(msg, vbYesNo + vbQuestion, "Query") = 6) Then
                    If CaseFileName = "Untitled" Then
                        FileSaveAs()
                    Else
                        SaveCurrentFile()
                    End If
                End If
            End If

            '   Clear the symptom selections.
            MustBox.Text = ""
            MustData = ""
            If NoSelectedItems = False Then ' This would cause an exception if data source was not already connected
                SelSympData.Clear()
                SelLst.DataSource = Nothing ' Remove and re-associate database in order to refresh SelLst
                NoSelectedItems = True
            End If

            If SymSelectedSize <> 0 Then
                ' re-initialize symSelected array
                For ptr = 0 To SymSelectedSize - 1
                    SymSelected(ptr) = -1
                Next
                SymSelectedSize = 0 ' Number of indices in SelectedIndicesArray
            End If

            Lst1.SelectedItems.Clear()
            CaseFileName = "Untitled"
            WriteTitles()     ' Display file names at top of screen.
            EnableFileClose()
        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub

    Public Sub FileOpen()
        Dim fileStatus As Boolean   ' true = good
        Dim firstLine As Boolean   'true indicates first file line
        Dim lineDat As String   'line read from file
        Dim locSympData As String   ' Symptom data from Lst1
        Dim locSympDesc As String   ' Symptom text from Lst1
        Dim locSympIDInt As Integer ' integer version of locSympIDStr
        Dim locWeightInt As Integer ' integer version of locWeightStr
        Dim locSympIDStr As String  ' Symptom ID read from file
        Dim locWeightStr As String  ' Weight read from file
        Dim msg As String
        Dim NewCaseFilePath As String   ' case file path returned from SaveFileDialog
        Dim spacePos As Integer   ' points to space in string
        Dim strPtr As Integer   ' string pointer
        Dim temp As String  ' portion of file line, used to determine version number
        Dim versionNumber As Integer    'number of DHA version, * 10

        Try
            ErrorFlag = 1007
            ' Delete temporary file.
            FileSystem.FileClose(CaseFile)
            OpenFileDialog1.FileName = ""

            ' Set Filters.
            OpenFileDialog1.Filter = "All Files (*.*)|*.*|Case Files (*.CSE)|*.CSE"

            ' Specify default filter.
            OpenFileDialog1.FilterIndex = 2
            CaseFilePath = (GetSetting("Hahnasst", "Startup", "CaseFilePath", [Default]:=AppContext.BaseDirectory))
            OpenFileDialog1.InitialDirectory = CaseFilePath

            ' Display the Open dialog box.
            If OpenFileDialog1.ShowDialog() = vbOK Then

                ' Call the open file procedure.
                If OpenFileDialog1.FileName <> "" Then  ' If a file was selected
                    Cursor = Cursors.WaitCursor    ' Change cursor to hourglass.
                    selectedFile = OpenFileDialog1.FileName
                    strPtr = Len(selectedFile)
                    ' Find location of last "\"
                    While ((strPtr > 0) And (Strings.Mid(selectedFile, strPtr, 1) <> "\"))
                        strPtr -= 1
                    End While
                    NewCaseFilePath = Strings.Left(selectedFile, strPtr)
                    CaseFileName = Strings.Mid(selectedFile, strPtr + 1)
                    If NewCaseFilePath <> CaseFilePath Then  ' Need to update path in registry
                        SaveSetting("Hahnasst", "Startup", "CaseFilePath", NewCaseFilePath)
                        CaseFilePath = NewCaseFilePath
                    End If
                    If FileLen(CaseFilePath + CaseFileName) > 0 Then
                        ErrorFlag = 1008

                        fileStatus = OpenFileRead(CaseFilePath + CaseFileName, CaseFile)
                        If Not fileStatus Then Throw New System.Exception("An exception has occurred.")    ' Force a trap

                        ErrorFlag = 1009
                        WriteTitles()   ' Write title with updated case file name to screen
                        firstLine = True
                        '   Clear the symptom selections.
                        MustBox.Text = ""
                        MustData = ""
                        If NoSelectedItems = False Then ' This would cause an exception if data source was not already connected
                            SelSympData.Clear()
                            SelLst.DataSource = Nothing ' Remove and re-associate database in order to refresh SelLst
                            NoSelectedItems = True
                        End If

                        Try
                            lineDat = LineInput(CaseFile)
                            If InStr(lineDat, " ") <> 0 Then    ' Version previous to 3.0
                                msg = "Case file is from a previous version; saving this file will update it to version 3.0."
                                Dim unused1 = MsgBox(Prompt:=msg,
                                Buttons:=vbOKOnly + vbInformation,
                                Title:="Information")
                                ' Version 2 has 1st line format:  Must Symp ID <space> Must Weight <space> version number (30 for 3.0); will have 2 spaces
                                ' Version 1 has 1st line format:  Must Symp ID <space> Must Weight (no version number); will have 1 space
                                temp = Strings.Mid(lineDat, InStr(lineDat, " ") + 1)  ' Skip past 1st space
                                If InStr(temp, " ") <> 0 Then
                                    temp = Strings.Mid(temp, InStr(temp, " ") + 1)    ' Skip past 2nd space in lineDat
                                    versionNumber = Val(temp)
                                Else
                                    versionNumber = 10
                                End If
                            Else    'Version 3.0 or later
                                versionNumber = Val(lineDat)
                            End If
                            If versionNumber < 40 Then
                                ' Next record will be for MustBox; update it if record is not "0 0"
                                ' if version 1.0 or 2, don't need to read first file line again.
                                If versionNumber >= 30 Then    ' need to read the next line, it will be for the MustBox
                                    lineDat = LineInput(CaseFile)
                                End If
                                If Val(Strings.Left(lineDat, InStr(lineDat, " "))) <> 0 Then   ' Entry is MustBox data
                                    spacePos = InStr(lineDat, " ")
                                    If spacePos = 0 Then Throw New System.Exception("An exception has occurred.")   ' corrupt file
                                    locSympIDStr = Strings.Left(lineDat, spacePos - 1)
                                    locWeightStr = Strings.Mid(lineDat, spacePos + 1)
                                    locSympIDInt = Val(locSympIDStr)
                                    locWeightInt = Val(locWeightStr)

                                    ' Get text and data from Lst1 item
                                    locSympDesc = Lst1.Items(MapSymIDToIndex(locSympIDInt)).SympDesc.ToString
                                    locSympDesc = locWeightStr + vbTab + locSympDesc    ' Add symptom weight
                                    locSympData = Lst1.Items(MapSymIDToIndex(locSympIDInt)).SympData.ToString

                                    ' Add these to MustBox
                                    MustBox.Text = locSympDesc
                                    MustData = locSympData
                                    DeleteMustButton.Enabled = True

                                    ' Set the Lst1 item to Selected
                                    Lst1.SetSelected(MapSymIDToIndex(locSympIDInt), True)
                                End If
                                While Not EOF(CaseFile)
                                    ' Get remaining records, populate SelLst.
                                    lineDat = LineInput(CaseFile)
                                    If Val(lineDat) <> 0 Then
                                        spacePos = InStr(lineDat, " ")
                                        If spacePos = 0 Then Throw New System.Exception("An exception has occurred.")   ' corrupt file
                                        locSympIDStr = Strings.Left(lineDat, spacePos - 1)
                                        locWeightStr = Strings.Mid(lineDat, spacePos + 1)
                                        locSympIDInt = Val(locSympIDStr)
                                        locWeightInt = Val(locWeightStr)

                                        ' Get text and data from Lst1 item
                                        locSympDesc = Lst1.Items(MapSymIDToIndex(locSympIDInt)).SympDesc.ToString
                                        locSympDesc = locWeightStr + vbTab + locSympDesc    ' Add symptom weight
                                        locSympData = Lst1.Items(MapSymIDToIndex(locSympIDInt)).SympData.ToString

                                        ' Add these to SelLst
                                        SelSympData.Add(New SelSympDat(locSympDesc, locSympData))
                                        NoSelectedItems = False

                                        ' Set the Lst1 item to Selected
                                        Lst1.SetSelected(MapSymIDToIndex(locSympIDInt), True)
                                        DeleteSelButton.Enabled = True
                                    End If
                                End While
                                SortAndUpdateSelLst()
                            Else
                                msg = "Case file is from a future version; please install version " + versionNumber.ToString + " of the software to read it."
                                Dim unused2 = MsgBox(Prompt:=msg,
                                Buttons:=vbOKOnly + vbInformation,
                                Title:="Information")
                            End If
                        Catch
                            msg = "Case file appears to be corrupt; unable to load it."
                            Dim unused3 = MsgBox(Prompt:=msg,
                                                Buttons:=vbOKOnly + vbInformation,
                                                Title:="Information")
                            Exit Sub
                        End Try

                        If Lst1.SelectedItems.Count <> -1 Then
                            PrescribeButton.Enabled = True  ' If at least one symptom is selected, enable the Prescribe Button.
                            ' Make it appear in SelLst box
                            SelLst.DataSource = SelSympData
                            SelLst.ValueMember = "SelSympData"
                            SelLst.DisplayMember = "SelSympDesc"
                            SelLst.ClearSelected()
                            NoSelectedItems = False
                        End If

                        SelectRemedy.Enabled = True
                        Dirty = False
                        DisableFileNew()
                        DisableFileOpen()
                        EnableFileSave()
                        EnableFileSaveAs()
                        EnableFileClose()
                        EnableFilePrintPreview()
                        EnableFilePrint()
                    Else    ' File doesn't contain any data
                        msg = "Selected case file doesn't contain any data."
                        Dim unused = MsgBox(Prompt:=msg,
                    Buttons:=vbOKOnly + vbInformation,
                            Title:="File Empty")
                    End If ' If "Cancel" wasn't pressed

                End If
            End If
            Cursor = Cursors.Default
            Exit Sub

        Catch
            Cursor = Cursors.Default

            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")

        End Try
    End Sub


    Public Sub FilePrint()
        ' Load the print buffer with Symptom list

        Dim msg As String           ' Error message string
        Dim Ptr As Integer          ' loop pointer
        Dim S As String             ' line from list box
        Dim StrPtr As Integer   ' string pointer

        Try
            ErrorFlag = 1014

            S = " "   ' Avoid warning variable used before assigning value
            Cursor = Cursors.WaitCursor    ' Change cursor to hourglass.
            If (MustBox.Text <> "") Then
                ErrorFlag = 1015
                S = MustBox.Text
                stringBuffer += "Must-use symptom:"
                StrPtr = InStr(S, "\")    ' Find start of ID's string
                If (StrPtr <> 0) Then
                    S = Mid(S, StrPtr) - 1
                    stringBuffer += Chr(9) + S + vbCrLf
                End If
            End If

            stringBuffer += "Symptom list:" + vbCrLf
            ErrorFlag = 1016

            For Ptr = 0 To SelLst.Items.Count - 1
                S = SelLst.Items(Ptr).SelSympDesc.ToString()
                If (S <> "") Then
                    stringBuffer += S + vbCrLf
                End If
            Next Ptr

            Cursor = Cursors.Default    ' Change cursor to default.
        Catch
            Cursor = Cursors.Default    ' Change cursor to default.

            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub

    Public Sub FileQuestClose()
        Dim msg As String   ' dialog box string

        Try
            ErrorFlag = 1017

            If (QDirty) Then ' Ask if user wants to save the file.
                msg = "Data has changed; do you want to save the current quest. file?"
                ' If data changed, allow user to save current file.
                If (MsgBox(msg, vbYesNo + vbQuestion) = 6) Then
                    If FirstQuestFileSave Then
                        Call FileQuestSaveAs()
                    Else
                        Call SaveCurrentQuestFile()
                    End If
                End If
            End If
            ClearQuestionaire()

            FileSystem.FileClose(QuestFile)

            EnableFileQuestNew()
            EnableFileQuestOpen()
            DisableFileQuestSave()
            DisableFileQuestSaveAs()
            DisableFileQuestClose()

            QuestFileName = "Untitled"
            WriteTitles()     ' Display file names at top of screen.

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub

    Public Sub FileQuestNew()

        Dim ErrStyle As Integer
        Dim msg As String

        Try
            ErrorFlag = 1018
            ErrStyle = vbOKOnly + vbCritical
            If SympSize = 0 Then
                SympSize = 1
            End If
            SymPtr = 0
            Questionnaire.Top = 0
            Questionnaire.Left = 0
            Questionnaire.Show()

            If (QDirty) Then  ' Ask if user wants to save the file.
                msg = "Questionaire data has changed; do you want to save the current questionaire file?"
                If (MsgBox(msg, vbYesNo + vbQuestion, "Query") = 6) Then
                    If QuestFileName = "Untitled" Then
                        FileQuestSaveAs()
                    Else
                        SaveCurrentQuestFile()
                    End If
                End If
            End If

            '   Clear the form.
            ClearQuestionaire()

            QuestFileName = "Untitled"
            WriteTitles()     ' Display file names at top of screen.
            EnableFileQuestClose()
            Exit Sub

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")

        End Try
    End Sub

    Public Sub FileQuestOpen()
        Dim FileStatus As Boolean   ' true = good
        Dim ItmPtr As Integer
        Dim LineDat As String   'line read from file
        Dim msg As String
        Dim NewQuestFilePath As String   ' case file path returned from SaveFileDialog
        Dim strPtr As Integer   ' string pointer
        Dim SympPtr As Integer

        Try
            ErrorFlag = 1021
            Cursor = Cursors.WaitCursor    ' Change cursor to hourglass.
            If SympSize = 0 Then
                SympSize = 1
            End If
            SymPtr = 0
            Questionnaire.Top = 0
            Questionnaire.Left = 0
            Questionnaire.Show()

            FileSystem.FileClose(QuestFile)
            OpenFileDialog1.FileName = QuestFileName
            ' Set Filters.
            OpenFileDialog1.Filter = "All Files (*.*)|*.*|Quest. Files (*.QUE)|*.QUE"
            ' Specify default filter.
            OpenFileDialog1.FilterIndex = 2
            QuestFilePath = (GetSetting("Hahnasst", "Startup", "QuestFilePath", [Default]:=AppContext.BaseDirectory))
            OpenFileDialog1.InitialDirectory = QuestFilePath

            ' Display the Open dialog box.
            If OpenFileDialog1.ShowDialog() = vbOK Then
                ' Call the open file procedure.
                If OpenFileDialog1.FileName <> "" Then  ' If a file was selected
                    Cursor = Cursors.WaitCursor    ' Change cursor to hourglass.
                    selectedFile = OpenFileDialog1.FileName
                    strPtr = Len(selectedFile)
                    ' Find location of last "\"
                    While ((strPtr > 0) And (Strings.Mid(selectedFile, strPtr, 1) <> "\"))
                        strPtr -= 1
                    End While
                    NewQuestFilePath = Strings.Left(selectedFile, strPtr)
                    QuestFileName = Strings.Mid(selectedFile, strPtr + 1)
                    If NewQuestFilePath <> QuestFilePath Then  ' Need to update path in registry
                        SaveSetting("Hahnasst", "Startup", "QuestFilePath", NewQuestFilePath)
                        QuestFilePath = NewQuestFilePath
                    End If
                    If FileLen(selectedFile) > 0 Then
                        ErrorFlag = 1020

                        FileStatus = OpenFileRead(QuestFilePath + QuestFileName, QuestFile)
                        If Not FileStatus Then Throw New System.Exception("An exception has occurred.")    ' Force a trap
                        WriteTitles()   ' Write title with updated case file name to screen

                        Try
                            LineDat = LineInput(QuestFile)
                            SympSize = Val(LineDat)
                            ReDim Symptom.SDate(SympSize - 1)
                            ReDim Symptom.Title(SympSize - 1)
                            ReDim Symptom.P1(6, SympSize - 1)
                            ReDim Symptom.P1Cb(4, SympSize - 1)
                            ReDim Symptom.P2(6, SympSize - 1)
                            ReDim Symptom.P3(1, SympSize - 1)
                            ReDim Symptom.P3Cb(49, SympSize - 1)
                            ReDim Symptom.P4Cb(4, SympSize - 1)
                            ReDim Symptom.P4(1, SympSize - 1)
                            ReDim Record.P5AgeCb(4)
                            ReDim Record.P5SexCb(1)
                            ReDim Record.P5(5)
                            ReDim Record.P6(5)
                            ReDim Record.P6Cb(14)
                            ReDim Record.P7(14)
                            ReDim Record.P7Cb(45)
                            For SympPtr = 0 To SympSize - 1
                                Symptom.SDate(SympPtr) = LineInput(QuestFile)
                                Symptom.Title(SympPtr) = LineInput(QuestFile)
                                For ItmPtr = 0 To 6
                                    Symptom.P1(ItmPtr, SympPtr) = LineInput(QuestFile)
                                Next ItmPtr
                                For ItmPtr = 0 To 4
                                    LineDat = LineInput(QuestFile)
                                    If Mid(LineDat, 1, 1) = "1" Then
                                        Symptom.P1Cb(ItmPtr, SympPtr) = True
                                    Else
                                        Symptom.P1Cb(ItmPtr, SympPtr) = False
                                    End If
                                Next ItmPtr
                                For ItmPtr = 0 To 6
                                    Symptom.P2(ItmPtr, SympPtr) = LineInput(QuestFile)
                                Next ItmPtr
                                For ItmPtr = 0 To 1
                                    Symptom.P3(ItmPtr, SympPtr) = LineInput(QuestFile)
                                Next ItmPtr
                                For ItmPtr = 0 To 49
                                    LineDat = LineInput(QuestFile)
                                    If Mid(LineDat, 1, 1) = "1" Then
                                        Symptom.P3Cb(ItmPtr, SympPtr) = True
                                    Else
                                        Symptom.P3Cb(ItmPtr, SympPtr) = False
                                    End If
                                Next ItmPtr
                                For ItmPtr = 0 To 4
                                    LineDat = LineInput(QuestFile)
                                    If Mid(LineDat, 1, 1) = "1" Then
                                        Symptom.P4Cb(ItmPtr, SympPtr) = True
                                    Else
                                        Symptom.P4Cb(ItmPtr, SympPtr) = False
                                    End If
                                Next ItmPtr
                                For ItmPtr = 0 To 1
                                    Symptom.P4(ItmPtr, SympPtr) = LineInput(QuestFile)
                                Next ItmPtr
                            Next SympPtr

                            Record.Name = LineInput(QuestFile)
                            Record.PDate = LineInput(QuestFile)
                            Record.DOB = LineInput(QuestFile)
                            For ItmPtr = 0 To 4
                                LineDat = LineInput(QuestFile)
                                If Mid(LineDat, 1, 1) = "1" Then
                                    Record.P5AgeCb(ItmPtr) = True
                                Else
                                    Record.P5AgeCb(ItmPtr) = False
                                End If
                            Next ItmPtr
                            For ItmPtr = 0 To 1
                                LineDat = LineInput(QuestFile)
                                If Mid(LineDat, 1, 1) = "1" Then
                                    Record.P5SexCb(ItmPtr) = True
                                Else
                                    Record.P5SexCb(ItmPtr) = False
                                End If
                            Next ItmPtr
                            For ItmPtr = 0 To 5
                                Record.P5(ItmPtr) = LineInput(QuestFile)
                            Next ItmPtr
                            For ItmPtr = 0 To 5
                                Record.P6(ItmPtr) = LineInput(QuestFile)
                            Next ItmPtr
                            For ItmPtr = 0 To 14
                                LineDat = LineInput(QuestFile)
                                If Mid(LineDat, 1, 1) = "1" Then
                                    Record.P6Cb(ItmPtr) = True
                                Else
                                    Record.P6Cb(ItmPtr) = False
                                End If
                            Next ItmPtr
                            For ItmPtr = 0 To 14
                                Record.P7(ItmPtr) = LineInput(QuestFile)
                            Next ItmPtr
                            For ItmPtr = 0 To 44
                                LineDat = LineInput(QuestFile)
                                If Mid(LineDat, 1, 1) = "1" Then
                                    Record.P7Cb(ItmPtr) = True
                                Else
                                    Record.P7Cb(ItmPtr) = False
                                End If
                            Next ItmPtr

                            DisableFileQuestNew()
                            DisableFileQuestOpen()
                            EnableFileQuestSave()
                            EnableFileQuestSaveAs()
                            EnableFileQuestClose()
                            EnableFilePrintPreview()
                            EnableFilePrint()
                            If SympSize > 1 Then
                                Questionnaire.EnableNextSymptom() ' Enable Next Symptom
                            End If
                        Catch
                            msg = "Selected questionnaire file appears to be corrupt; unable to load it."
                            Dim unused3 = MsgBox(Prompt:=msg,
                        Buttons:=vbOKOnly + vbInformation,
                        Title:="Information")
                            Exit Sub
                        End Try
                    Else
                        ' File doesn't contain any data
                        msg = "Selected questionnaire file doesn't contain any data."
                        Dim unused = MsgBox(Prompt:=msg,
                    Buttons:=vbOKOnly + vbInformation,
                            Title:="File Empty")
                    End If
                End If
            End If
            Cursor = Cursors.Default
            Exit Sub

        Catch
            Cursor = Cursors.Default
            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")

            Exit Sub
        End Try
    End Sub

    Public Sub FileQuestSaveAs()
        Dim msg As String   ' error message string
        Dim myStream As Stream = Nothing
        Dim NewQuestFilePath As String   ' questionnaire file path returned from SaveFileDialog
        Dim strPtr As Integer   ' string pointer
        Dim OldQuestFileName As String   ' previous value of CaseFileName
        Dim WrFile As Boolean   ' Write file permissive

        Try
            ErrorFlag = 1021
            QuestFilePath = (GetSetting("Hahnasst", "Startup", "QuestFilePath", [Default]:=AppContext.BaseDirectory))
            ' Set Filters.
            SaveFileDialog1.InitialDirectory = "c:\"
            SaveFileDialog1.Filter = "Quest. Files (*.QUE) |*.QUE| All Files (*.*) | *.*"
            ' Specify default filter.
            SaveFileDialog1.FilterIndex = 1
            SaveFileDialog1.InitialDirectory = QuestFilePath

            OldQuestFileName = QuestFileName
            WrFile = False

            ' Display the Save dialog box.
            SaveFileDialog1.InitialDirectory = QuestFilePath
            SaveFileDialog1.FileName = QuestFileName
            If SaveFileDialog1.ShowDialog() = vbOK Then

                ' If the file name is not an empty string open it for saving.
                If SaveFileDialog1.FileName <> "" Then
                    selectedFile = SaveFileDialog1.FileName
                    strPtr = Len(selectedFile)
                    ' Find location of last "\"
                    While ((strPtr > 0) And (Strings.Mid(selectedFile, strPtr, 1) <> "\"))
                        strPtr -= 1
                    End While
                    NewQuestFilePath = Strings.Left(selectedFile, strPtr)
                    WrFile = True
                    QuestFileName = Strings.Mid(selectedFile, strPtr + 1)
                    If NewQuestFilePath <> QuestFilePath Then  ' Need to update path in registry
                        SaveSetting("Hahnasst", "Startup", "QuestFilePath", NewQuestFilePath)
                        QuestFilePath = NewQuestFilePath
                    End If
                End If
                ErrorFlag = 1022
                If WrFile Then
                    FileSystem.FileClose(QuestFile)
                    Call SaveQFile(QuestFilePath + QuestFileName)
                End If
                WriteTitles()     ' Display file names at top of screen.
            End If

            QDirty = False

        Catch Ex As Exception
            MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                            Buttons:=vbOKOnly + vbCritical,
                            Title:="ERROR!")
        End Try

    End Sub

    Public Function OpenFileReadRandom(FileName As String, FileNumber As Integer, RecLen As Integer)
        Dim FileStatus As Boolean
        FileStatus = True   ' Assume success.
        Err.Description = ""
        FileSystem.FileOpen(FileNumber, FileName, Mode:=OpenMode.Random, RecordLength:=RecLen)
        If Err.Description = "File already open" Then
            FileSystem.FileClose(FileNumber)
            Err.Description = ""
            FileSystem.FileOpen(FileNumber, FileName, Mode:=OpenMode.Random, RecordLength:=RecLen)
            If Err.Description = "File Already Open" Then FileStatus = False
        End If
        Return (FileStatus)
    End Function

    Public Function GetAutoweight(Symptom As String) As Integer
        Dim AFileWord As String ' AutoWt.dat word
        Dim AW As Integer   ' weight determined for Symptom
        Dim AWStr As String ' string form of AW
        Dim DONE As Boolean     ' weight was found flag
        Dim First As Integer    ' top of binary search
        Dim Found As Boolean    ' word was found flag
        Dim Last As Integer     ' bottom of binary search
        Dim loc_Afile_pos As Integer    ' local search file position
        Dim locLineDat As String    ' line read from AutoWt.dat
        Dim Ptr1 As Integer   ' string pointer
        Dim SympWord As String  ' word being looked up

        Try
            Dim AutoWtText() = System.IO.File.ReadAllLines(AppContext.BaseDirectory + DATA_DIR + "AutoWt.dat")

            ' Assign AW a value of 1 for physical symptoms, 2 for general symptoms, or 3 for mind symptoms; leave 0 if unable to classify.
            RaiseException = False
            AW = 0
            Afile_size = Val(System.IO.File.ReadAllLines(AppContext.BaseDirectory + DATA_DIR + "AutoWt.dat").Length.ToString())

            '   Do a binary search of AutoWt.dat for words in the string.
            DONE = False
            Ptr1 = 0
            While (Not DONE Or AW = 2) And Ptr1 < Len(Symptom)
                Ptr1 = InStr(Symptom, " ") - 1
                If Ptr1 = -1 Then Ptr1 = Len(Symptom)
                SympWord = LCase(Mid(Symptom, 1, Ptr1))
                If Ptr1 <> Len(Symptom) Then Symptom = Mid(Symptom, Ptr1 + 2)
                If Mid(SympWord, Len(SympWord), 1) = "," Or
                    Mid(SympWord, Len(SympWord), 1) = ";" Or
                    Mid(SympWord, Len(SympWord), 1) = "." Or
                    Mid(SympWord, Len(SympWord), 1) = ")" Then _
               SympWord = Mid(SympWord, 1, Len(SympWord) - 1)
                Last = Afile_size
                First = 0

                ' Find text position in file.
                Found = False
                AW = 0
                While Not Found
                    loc_Afile_pos = (First + Last + 0.1) / 2
                    If (loc_Afile_pos < 1) Then loc_Afile_pos = 1
                    If (loc_Afile_pos >= Afile_size) Then _
                    loc_Afile_pos = Afile_size - 1
                    locLineDat = AutoWtText(loc_Afile_pos)

                    If Last - First <= 1 Then
                        If (First > 0) Then
                            loc_Afile_pos = First
                            locLineDat = AutoWtText(loc_Afile_pos)
                            AFileWord = Mid(locLineDat, 1, InStr(locLineDat, " ") - 1)
                            If SympWord <> AFileWord Then _
                            loc_Afile_pos = Last
                        End If
                        Found = True
                        AWStr = Mid(locLineDat, InStr(locLineDat, " ") + 1, 1)
                        AW = CInt(AWStr)
                    End If
                    AFileWord = Mid(locLineDat, 1, InStr(locLineDat, " ") - 1)
                    If SympWord = AFileWord Then
                        Found = True
                        AWStr = Mid(locLineDat, InStr(locLineDat, " ") + 1, 1)
                        AW = CInt(AWStr)
                    ElseIf SympWord > AFileWord Then
                        First = (First + Last + 0.1) / 2
                    Else
                        Last = (First + Last + 0.1) / 2
                    End If
                End While
                If AW <> 0 Then DONE = True
            End While
            GetAutoweight = AW
            Exit Function

        Catch
            GetAutoweight = -1
            RaiseException = True
            Exit Function
        End Try
    End Function

    Public Sub GetQFileName()
        Dim msg As String

        msg = "Questionnaire has been changed.  Do you wish to save your changes?"
        If (MsgBox(msg, vbYesNo + vbQuestion) = 6) Then
            SaveQuestFile.Select()
        End If
    End Sub

    Public Sub OptionsAutoWeight()
        If Holostic.Checked Then
            AutoWeight = True
        Else
            AutoWeight = True
        End If

    End Sub


    Public Sub OptionsNormalize()
        If boxNormalize.Checked Then
            Normalize = True
        Else
            Normalize = False
        End If

    End Sub


    Public Sub OptionsSetPreferences()
        SetPref.Show()
    End Sub

    Public Sub FileSaveAs()
        Dim FileStatus As Boolean   ' true = good status
        Dim msg As String       ' message dialog text
        Dim NewCaseFilePath As String   ' case file path returned from SaveFileDialog
        Dim strPtr As Integer   ' string pointer
        Dim WrFile As Boolean   ' Write file permissive

        Try
            ErrorFlag = 1023
            CaseFilePath = (GetSetting("Hahnasst", "Startup", "CaseFilePath", [Default]:=".\"))
            ' Set Filters.
            SaveFileDialog1.Filter = "Case Files (*.CSE) |*.CSE| All Files (*.*) | *.*"
            ' Specify default filter.
            SaveFileDialog1.FilterIndex = 1
            WrFile = False

            ' Display the Save dialog box.
            SaveFileDialog1.InitialDirectory = CaseFilePath
            SaveFileDialog1.FileName = CaseFileName
            If (SaveFileDialog1.ShowDialog() = vbOK) Then
                ' Call the save file procedure.
                If SaveFileDialog1.FileName <> "" Then    'get the file information
                    selectedFile = SaveFileDialog1.FileName
                    strPtr = 0

                    strPtr = Len(selectedFile)
                    ' Find location of last "\"
                    While ((strPtr > 0) And (Strings.Mid(selectedFile, strPtr, 1) <> "\"))
                        strPtr -= 1
                    End While
                    NewCaseFilePath = Strings.Left(selectedFile, strPtr)
                    WrFile = True
                    CaseFileName = Strings.Mid(selectedFile, strPtr + 1)
                    If NewCaseFilePath <> CaseFilePath Then  ' Need to update path in registry
                        SaveSetting("Hahnasst", "Startup", "CaseFilePath", NewCaseFilePath)
                        CaseFilePath = NewCaseFilePath
                    End If
                End If

                If (WrFile) Then
                    ErrorFlag = 1024
                    FileSystem.FileClose(CaseFile)

                    FileStatus = OpenFileWrite(CaseFilePath + CaseFileName, CaseFile)
                    Call SaveCurrentFile()
                End If
                WriteTitles()     ' Display file names at top of screen.
            End If

            Dirty = False

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try

    End Sub


    Public Sub OptionsDefault()
        SelectionMode = 1
    End Sub

    Public Sub OptionsEqualWeight()
        SelectionMode = 2

    End Sub


    Public Sub OptionsStraightCount()
        SelectionMode = 3
    End Sub

    Public Sub QuestPager(ListID As Object)
        Dim Temp As String

        Temp = Mid(ListID.List(ListID.ListIndex), 6, 1)
        Select Case Temp
            Case "1"
                Questionnaire.TabPage1.Show()
            Case "2"
                Questionnaire.TabPage2.Show()
            Case "3"
                Questionnaire.TabPage3.Show()
            Case "4"
                Questionnaire.TabPage4.Show()
            Case "5"
                Questionnaire.TabPage5.Show()
            Case "6"
                Questionnaire.TabPage6.Show()
            Case "7"
                Questionnaire.TabPage7.Show()
        End Select
    End Sub

    Public Sub SaveCurrentQuestFile()
        Dim msg As String       ' error message string

        '   Handle writing current questionaire data to common.
        Try
            ErrorFlag = 1025

            SaveQFile(QuestFilePath + QuestFileName)
            QDirty = False
            FirstQuestFileSave = False
            Exit Sub

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
            Exit Sub
        End Try
    End Sub

    Public Sub SaveCurrentFile()
        Dim commaPos As Integer ' position in string of first comma
        Dim filedat As String   ' line of data for case file
        Dim FileStatus As Boolean   ' true = good
        Dim msg As String   ' Message box text
        Dim selPtr As Integer   ' select list pointer
        Dim locSelSympData As String   ' data from symptom in SelLst
        Dim locSelSympDesc As String   ' text from item in SelLst
        Dim strSympID As String ' string version of symptom ID
        Dim strSympWeight As String ' string version of symptom weight

        ' First line in file is version number.
        ' Second line is MustBox symptom ID and Weight, or 0's if nothing in MustBox
        ' Subsequent lines are SelLst items symptom ID and Weight, if SelLst isn't empty
        Try
            ErrorFlag = 1026
            FileSystem.FileClose(CaseFile)  ' In case file was opened as "Read"
            FileStatus = OpenFileWrite(CaseFilePath + CaseFileName, CaseFile)
            FileSystem.Print(CaseFile, DHA_Version) ' First file entry is version number
            FileSystem.Print(CaseFile, vbCrLf)
            ' If MustBox contains a symptom, write its symptom ID and weight, otherwise write 0's as placeholders
            If MustBox.Text <> "" Then
                commaPos = InStr(MustData, ",")
                strSympID = Strings.Left(MustData, commaPos - 1)
                strSympWeight = Strings.Left(MustBox.Text, 1)
            Else
                strSympID = "0"
                strSympWeight = "0"
            End If
            filedat = strSympID + " " + strSympWeight + vbCrLf
            FileSystem.Print(CaseFile, filedat)

            For selPtr = 0 To SelLst.Items.Count - 1
                locSelSympData = SelLst.Items(selPtr).SelSympData.ToString
                commaPos = InStr(locSelSympData, ",")
                strSympID = Strings.Left(locSelSympData, commaPos - 1)
                locSelSympDesc = SelLst.Items(selPtr).selSympDesc.ToString
                strSympWeight = Strings.Left(locSelSympDesc, 1)
                filedat = strSympID + " " + strSympWeight + vbCrLf
                FileSystem.Print(CaseFile, filedat)
            Next selPtr
            Dirty = False

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub

    Sub SaveQFile(QuestFileName)
        Dim FileStatus
        Dim msg As String   ' error message string

        Try
            ErrorFlag = 1060
            FileSystem.FileClose(QuestFile)  ' In case file was opened as "Read"
            FileStatus = OpenFileWrite(QuestFileName, QuestFile)
            FileSystem.Print(QuestFile, Str(SympSize) + vbCrLf)

            For SympPtr = 0 To SympSize - 1
                If Symptom.SDate(SymPtr) = "" Then Symptom.SDate(SymPtr) = Questionnaire.DateTimePicker1.Value    ' Date was never changed, need to update with date from screen.
                FileSystem.Print(QuestFile, Symptom.SDate(SympPtr) + vbCrLf)
                FileSystem.Print(QuestFile, Symptom.Title(SympPtr) + vbCrLf)
                For ItmPtr = 0 To 6
                    FileSystem.Print(QuestFile, Symptom.P1(ItmPtr, SympPtr) + vbCrLf)
                Next ItmPtr
                For ItmPtr = 0 To 4
                    If Symptom.P1Cb(ItmPtr, SympPtr) Then
                        FileSystem.Print(QuestFile, "1" + vbCrLf)
                    Else
                        FileSystem.Print(QuestFile, "0" + vbCrLf)
                    End If
                Next ItmPtr
                For ItmPtr = 0 To 6
                    FileSystem.Print(QuestFile, Symptom.P2(ItmPtr, SympPtr) + vbCrLf)
                Next ItmPtr
                For ItmPtr = 0 To 1
                    FileSystem.Print(QuestFile, Symptom.P3(ItmPtr, SympPtr) + vbCrLf)
                Next ItmPtr
                For ItmPtr = 0 To 49
                    If Symptom.P3Cb(ItmPtr, SympPtr) Then
                        FileSystem.Print(QuestFile, "1" + vbCrLf)
                    Else
                        FileSystem.Print(QuestFile, "0" + vbCrLf)
                    End If
                Next ItmPtr
                For ItmPtr = 0 To 4
                    If Symptom.P4Cb(ItmPtr, SympPtr) Then
                        FileSystem.Print(QuestFile, "1" + vbCrLf)
                    Else
                        FileSystem.Print(QuestFile, "0" + vbCrLf)
                    End If
                Next ItmPtr
                For ItmPtr = 0 To 1
                    FileSystem.Print(QuestFile, Symptom.P4(ItmPtr, SympPtr) + vbCrLf)
                Next ItmPtr
            Next SympPtr

            FileSystem.Print(QuestFile, Record.Name + vbCrLf)
            If Record.PDate = "" Then Record.PDate = Questionnaire.DateTimePicker1.Value
            FileSystem.Print(QuestFile, Record.PDate + vbCrLf)
            If Record.DOB = "" Then Record.DOB = Questionnaire.DateTimePicker2.Value
            FileSystem.Print(QuestFile, Record.DOB + vbCrLf)
            For ItmPtr = 0 To 4
                If Record.P5AgeCb(ItmPtr) Then
                    FileSystem.Print(QuestFile, "1" + vbCrLf)
                Else
                    FileSystem.Print(QuestFile, "0" + vbCrLf)
                End If
            Next ItmPtr
            For ItmPtr = 0 To 1
                If Record.P5SexCb(ItmPtr) Then
                    FileSystem.Print(QuestFile, "1" + vbCrLf)
                Else
                    FileSystem.Print(QuestFile, "0" + vbCrLf)
                End If
            Next ItmPtr
            For ItmPtr = 0 To 5
                FileSystem.Print(QuestFile, Record.P5(ItmPtr) + vbCrLf)
            Next ItmPtr
            For ItmPtr = 0 To 5
                FileSystem.Print(QuestFile, Record.P6(ItmPtr) + vbCrLf)
            Next ItmPtr
            For ItmPtr = 0 To 14
                If Record.P6Cb(ItmPtr) Then
                    FileSystem.Print(QuestFile, "1" + vbCrLf)
                Else
                    FileSystem.Print(QuestFile, "0" + vbCrLf)
                End If
            Next ItmPtr
            For ItmPtr = 0 To 14
                FileSystem.Print(QuestFile, Record.P7(ItmPtr) + vbCrLf)
            Next ItmPtr
            For ItmPtr = 0 To 44
                If Record.P7Cb(ItmPtr) Then
                    FileSystem.Print(QuestFile, "1" + vbCrLf)
                Else
                    FileSystem.Print(QuestFile, "0" + vbCrLf)
                End If
            Next ItmPtr
        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub
    '    Public Function UpdateRemList(id As String, MustUse As Boolean)
    Public Function UpdateRemList(LocRemID As Integer, MustUse As Boolean) As Boolean
        Dim LocMaxRemhits As Integer    ' local cound of max remedy hits
        Dim msg As String               ' error message string
        Dim Ptr1, Ptr2 As Integer       ' loop pointers

        Try
            ErrorFlag = 1039
            If NumRem > UBound(RemList) Then ReDim Preserve RemList(NumRem)
            RemList(NumRem) = LocRemID
            NumRem += 1

            ' Find max. number of hits.
            If (MustUse) Then
                LocMaxRemhits = 1
                If (LocMaxRemhits > MaxRemHits) Then _
                    MaxRemHits = LocMaxRemhits
            Else
                For Ptr1 = 0 To SelLst.Items.Count - 1
                    LocMaxRemhits = 1
                    For Ptr2 = Ptr1 To SelLst.Items.Count - 1
                        If (Ptr1 <> Ptr2) Then
                            If (RemList(Ptr1) = RemList(Ptr2)) Then _
                        LocMaxRemhits = LocMaxRemhits + 1
                        End If
                    Next Ptr2
                    If (LocMaxRemhits > MaxRemHits) Then _
                MaxRemHits = LocMaxRemhits
                Next Ptr1
            End If
            If (MaxRemHits >= MinNumb) Then
                UpdateRemList = True
            Else
                UpdateRemList = False
            End If

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
            UpdateRemList = False
        End Try
    End Function

    Public Sub WriteQ1Screen()
        Dim ItmPtr As Integer    'pointer to page item

        If Symptom.SDate(SymPtr) = "" Then ' Date was never changed, need to read it from screen.
            Symptom.SDate(SymPtr) = Questionnaire.DateTimePicker1.Value
        End If
        If Symptom.SDate(SymPtr) <> "" Then
            Questionnaire.DateTimePicker3.Value = Symptom.SDate(SymPtr)
        Else
            Questionnaire.DateTimePicker3.Value = Date.FromOADate(0)    ' Use a "default" date to prevent exception
        End If
        Questionnaire.Label85.Text = SymPtr.ToString
        Questionnaire.Text1(1).Text = Symptom.Title(SymPtr)
        Questionnaire.Text1(3).Text = Symptom.P1(0, SymPtr)
        Questionnaire.Text1(4).Text = Symptom.P1(1, SymPtr)
        Questionnaire.Text1(5).Text = Symptom.P1(2, SymPtr)
        Questionnaire.Text1(6).Text = Symptom.P1(3, SymPtr)
        Questionnaire.Text1(7).Text = Symptom.P1(4, SymPtr)
        Questionnaire.Text1(8).Text = Symptom.P1(5, SymPtr)
        Questionnaire.Text1(9).Text = Symptom.P1(6, SymPtr)
        For ItmPtr = 0 To 4
            If Symptom.P1Cb(ItmPtr, SymPtr) = True Then
                Questionnaire.Check1(ItmPtr).Checked = True
            Else
                Questionnaire.Check1(ItmPtr).Checked = False
            End If
        Next

        Questionnaire.Text1(0).Text = Record.Name
        If Record.PDate <> "" Then
            Questionnaire.DateTimePicker1.Value = Record.PDate
        Else
            Questionnaire.DateTimePicker1.Value = Date.Now
        End If

        If Record.DOB <> "" Then
            Questionnaire.DateTimePicker2.Value = Record.DOB
        Else
            Questionnaire.DateTimePicker2.Value = Date.Now
        End If

        If SymPtr = 0 Then
            Questionnaire.Command10.Enabled = False
        Else
            Questionnaire.Command10.Enabled = True
        End If
        If SymPtr = SympSize - 1 Then
            Questionnaire.Command11.Enabled = False
        Else
            Questionnaire.Command11.Enabled = True
        End If

    End Sub

    Public Sub WriteQ2Screen()
        Dim ItmPtr As Integer    'pointer to page item

        For ItmPtr = 0 To 6
            Questionnaire.Text1(ItmPtr + 10).Text = Symptom.P2(ItmPtr, SymPtr)
        Next ItmPtr

        If SymPtr = 0 Then
            Questionnaire.Command10.Enabled = False
        Else
            Questionnaire.Command10.Enabled = True
        End If
        If SymPtr = SympSize - 1 Then
            Questionnaire.Command11.Enabled = False
        Else
            Questionnaire.Command11.Enabled = True
        End If

    End Sub

    Public Sub WriteQ3Screen()
        Dim ItmPtr As Integer   'pointer to page item

        Questionnaire.Text1(17).Text = Symptom.P3(0, SymPtr)
        Questionnaire.Text1(18).Text = Symptom.P3(1, SymPtr)
        For ItmPtr = 0 To 49
            Questionnaire.Check1(ItmPtr + 5).Checked = Symptom.P3Cb(ItmPtr, SymPtr)
        Next

        If SymPtr = 0 Then
            Questionnaire.Command14.Enabled = False
        Else
            Questionnaire.Command14.Enabled = True
        End If
        If SymPtr = SympSize - 1 Then
            Questionnaire.Command15.Enabled = False
        Else
            Questionnaire.Command15.Enabled = True
        End If

    End Sub

    Public Sub WriteQ4Screen()
        Dim ItmPtr As Integer   'pointer to page item

        For ItmPtr = 0 To 3
            Questionnaire.Check1(ItmPtr + 55).Checked = Symptom.P4Cb(ItmPtr, SymPtr)
        Next
        Questionnaire.Text1(19).Text = Symptom.P4(0, SymPtr)
        Questionnaire.Text1(20).Text = Symptom.P4(1, SymPtr)
        WriteTitles()     ' Display file names, symptom numbers at top of screen.
        If SymPtr = 0 Then
            Questionnaire.Command18.Enabled = False
        Else
            Questionnaire.Command18.Enabled = True
        End If
        If SymPtr = SympSize - 1 Then
            Questionnaire.Command19.Enabled = False
        Else
            Questionnaire.Command19.Enabled = True
        End If
    End Sub

    Public Sub WriteQ5Screen()
        Dim ItmPtr As Integer    'pointer to page item

        For ItmPtr = 0 To 4
            Questionnaire.Check1(ItmPtr + 59).Checked = Record.P5AgeCb(ItmPtr)
        Next ItmPtr
        For ItmPtr = 0 To 1
            Questionnaire.Check1(ItmPtr + 64).Checked = Record.P5SexCb(ItmPtr)
        Next

        For ItmPtr = 0 To 5
            Questionnaire.Text1(ItmPtr + 21).Text = Record.P5(ItmPtr)
        Next ItmPtr

    End Sub

    Public Sub WriteQ6Screen()
        Dim ChkPtr As Integer   ' pointer to check box
        Dim ItmPtr As Integer   ' pointer to page item

        For ItmPtr = 0 To 5
            Questionnaire.Text1(ItmPtr + 27).Text = Record.P6(ItmPtr)
        Next ItmPtr
        For ItmPtr = 0 To 14
            Questionnaire.Check1(ItmPtr + 66).Checked = Record.P6Cb(ItmPtr)
        Next ItmPtr

        ChkPtr = 65
        For ItmPtr = 28 To 32
            If Questionnaire.Text1(ItmPtr).Text <> "" Then
                Questionnaire.Check1(ChkPtr).Enabled = True
                Questionnaire.Check1(ChkPtr + 1).Enabled = True
                Questionnaire.Check1(ChkPtr + 2).Enabled = True
            Else
                Questionnaire.Check1(ChkPtr).Enabled = False
                Questionnaire.Check1(ChkPtr + 1).Enabled = False
                Questionnaire.Check1(ChkPtr + 2).Enabled = False
            End If
            ChkPtr += 3
        Next
    End Sub

    Public Sub WriteQ7Screen()
        Dim ItmPtr As Integer  'pointer to page item
        Dim ChkPtr As Integer  ' Check box index

        For ItmPtr = 0 To 14
            Questionnaire.Text1(ItmPtr + 33).Text = Record.P7(ItmPtr)
        Next ItmPtr

        For ItmPtr = 0 To 44
            Questionnaire.Check1(ItmPtr + 81).Checked = Record.P7Cb(ItmPtr)
        Next ItmPtr

        ChkPtr = 81
        For ItmPtr = 33 To 47
            If Questionnaire.Text1(ItmPtr).Text <> "" Then
                Questionnaire.Check1(ChkPtr).Enabled = True
                Questionnaire.Check1(ChkPtr + 1).Enabled = True
                Questionnaire.Check1(ChkPtr + 2).Enabled = True
            Else
                Questionnaire.Check1(ChkPtr).Enabled = False
                Questionnaire.Check1(ChkPtr + 1).Enabled = False
                Questionnaire.Check1(ChkPtr + 2).Enabled = False
            End If
        Next
    End Sub

    Public Sub WriteTitles()
        Dim cFileName As String ' case file name only, without the path
        Dim ptr As Integer
        Dim qFileName As String ' questionnaire file name only, without the path

        ptr = InStrRev(CaseFileName, "\")
        If ptr > 0 Then
            cFileName = Strings.Mid(CaseFileName, ptr + 1)
        Else
            cFileName = CaseFileName
        End If

        ptr = InStrRev(QuestFileName, "\")
        If ptr > 0 Then
            qFileName = Strings.Mid(QuestFileName, ptr + 1)
        Else
            qFileName = QuestFileName
        End If

        Me.Text = "Dr. Hahnemann's Assistant - CASE FILE:  " + cFileName + " - QUEST FILE:  " + qFileName
        Questionnaire.Label83.Text = "Symptom " + Str(SymPtr + 1) + "  QUEST FILE:  " + qFileName
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim colorName As String     ' color name read from retistry
        Dim FileFound As Boolean    ' windows help file has been found flag
        Dim InName As String        ' name of symptom file for Lst1
        Dim Ptr, ptr1, ptr2 As Integer          ' loop counters
        Dim SymptomID As Integer    ' symptom ID in Lst1 records
        Dim LineCount As Integer    ' Number of lines in input file

        FirstLoad = True    ' Tells MouseMove handler to de-select ListBox1 item 0
        FirstRepSearch = True   ' Only want to run init items once when loading
        InLoad = True
        NoSelectedItems = True  ' Tells code to set SelLst data source when first item is added to listbox
        LastSelectedDataString = "" ' Used to supply Lst1 item data to "Must Use" box
        FindLoaded = False      ' Tells find_load() to execute its code only the first time it gets executed
        QuestFirstLoad = True   ' Tells questionnaire to start a new file and initialize all items to blanks when first loaded
        FileExitCalled = False  ' FileExitCalled prevents FileExit running more than once when Form1 is closed.
        CtrlPressed = False ' <CTRL> key is not currently pressed
        ErrorFlag = 1127

        ' Initialize print selections for main menu print
        CasePrint = False
        SympPrint = False
        RemPrint = False
        QPrint = False

        ' Initialize SymSelected array
        For Ptr = 0 To SymSelected.Count - 1
            SymSelected(Ptr) = -1
        Next

        '   Find windows help directory.ListSize
        FileFound = False
        QuestionnaireVisible = False    ' Questionnaire has not been loaded, so don't try to read its data.
        ptr1 = 2
        ShowSimilarSymptoms.Enabled = False
        RaiseException = False

        '  Load default settings.
        ' HKEY_USERS\S-1-5-21-616849596-177198723-2971636422-1001\SOFTWARE\VB and VBA Program Settings\
        ErrorFlag = 1128
        colorName = (GetSetting("Hahnasst", "Startup", "BackColor", [Default]:="Window"))
        If Strings.Left(colorName, 7) = "Color [" Then  ' Need to remove "Color" and brackets from registry engry
            colorName = Mid(colorName, 8)
            colorName = Strings.Left(colorName, Len(colorName) - 1)
        End If
        ScreenBackColor = Color.FromName(colorName)

        TextBox1.BackColor = ScreenBackColor
        Lst1.BackColor = ScreenBackColor
        MustBox.BackColor = ScreenBackColor
        SelLst.BackColor = ScreenBackColor

        colorName = (GetSetting("Hahnasst", "Startup", "ForeColor", [Default]:="WindowText"))
        If Strings.Left(colorName, 7) = "Color [" Then  ' Need to remove "Color" and brackets from registry engry
            colorName = Mid(colorName, 8)
            colorName = Strings.Left(colorName, Len(colorName) - 1)
        End If
        ScreenForeColor = Color.FromName(colorName)
        TextBox1.ForeColor = ScreenForeColor
        Lst1.ForeColor = ScreenForeColor
        MustBox.ForeColor = ScreenForeColor
        SelLst.ForeColor = ScreenForeColor

        ScreenFontName = New Font(GetSetting("Hahnasst", "Startup", "ScreenFontName", [Default]:="Microsoft Sans Serif"),
            GetSetting("Hahnasst", "Startup", "ScreenFontSize", [Default]:="7.8"))

        TextBox1.Font = ScreenFontName
        Lst1.Font = ScreenFontName
        MustBox.Font = ScreenFontName
        SelLst.Font = ScreenFontName

        PrinterFontName = New Font(GetSetting("Hahnasst", "Startup", "PrinterFontName", [Default]:="Microsoft Sans Serif"),
            GetSetting("Hahnasst", "Startup", "PrinterFontSize", [Default]:="8"))

        ShowToolbar = GetSetting("Hahnasst", "Startup",
        "ShowToolbar", True)
        If ShowToolbar Then
            ToolStrip1.Show()
        Else
            ToolStrip1.Hide()
        End If

        ShowStatusbar = GetSetting("Hahnasst", "Startup",
        "ShowStatusbar", False)
        If ShowStatusbar Then
            StatusStrip1.Show()
        Else
            StatusStrip1.Hide()
        End If

        ToolText = GetSetting("Hahnasst", "Startup",
        "ToolText", True)
        If ToolText = True Then
            TurnOnToolTips()
        Else
            TurnOffToolTips()
        End If

        ErrorFlag = 1129
        Me.Left = 0
        Me.Top = 0

        ShowDisclaimer = GetSetting("Hahnasst", "Startup",
        "ShowDisclaimer", True)
        ShowIntroduction = GetSetting("Hahnasst", "Startup",
        "ShowIntroduction", True)

        HideDisclaimer()
        HideIntroduction()

        If ShowDisclaimer Then
            VisibleDisclaimer()
        ElseIf ShowIntroduction Then
            VisibleIntroduction()
            ' Else - nothing to show
        End If

        SelectionMode = 1   ' Set to "Default"
        MinNumb = 1     ' Minimum number hits per remedy
        RepertorySearch = False
        FindAllVisible = False
        buttonDefault.Checked = True ' Default
        AutoWeight = False
        boxNormalize.Checked = True
        Normalize = True
        ExitPressed = False ' Forces exit if intro form EXIT button pressed.
        OptionsNormalize()
        SympSize = 0    ' Size of questionaire symptom array.
        SymSelectedSize = 0 ' Number of indices in SymSelected

        EnableFileNew()   ' File/New
        EnableFileOpen()   ' File/Open
        DisableFileSave()  ' File/Save
        DisableFileSaveAs()  ' File/SaveAs
        DisableFileClose()  ' File/Close
        DisableFilePrintPreview()   ' File/PrintPreview
        DisableFilePrint()  ' File/Print

        EnableFileQuestNew()  ' Quest. File/New
        EnableFileQuestOpen() ' Quest. File/Open
        DisableFileQuestSave()    ' Quest. File/Save
        DisableFileQuestSaveAs()  ' Quest. File/SaveAs
        DisableFileQuestClose()   ' Quest. File/Close

        SelectRemedy.Enabled = False    ' Disable Select Remedy Button
        MustUse.Enabled = False         ' Disable Must Use button

        ' Initialize Grid Search variable
        For ptr1 = 0 To 23
            GridString(ptr1) = ""
        Next

        ' Initialize combo box values
        Weight.Text = "1"
        MinNum.Text = "1"

        ' Load data files.
        Try
            InName = AppContext.BaseDirectory + DATA_DIR + "symptextdata.dat"
            LineCount = Val(System.IO.File.ReadAllLines(InName).Length.ToString())
            Dim SymptomData() = System.IO.File.ReadAllLines(AppContext.BaseDirectory + DATA_DIR + "symptextdata.dat")
            For ptr1 = 0 To LineCount - 1 Step 2   ' This took 10 - 11 seconds
                CombSympData.Add(New CombSympDat(SymptomData(ptr1), SymptomData(ptr1 + 1)))
            Next
            Lst1.DataSource = CombSympData
            Lst1.DisplayMember = "SympDesc"
            Lst1.ValueMember = "SympData"
            Lst1.ClearSelected()
        Catch
            Dim unused2 = MsgBox(Prompt:="symptextdata.dat file missing Or corrupted; please re-install Dr. Hahnemann's Assistant to fix this problem",
        Buttons:=vbOKOnly + vbCritical,
            Title:="ERROR!")
            Close()
        End Try

        Try
            Find.dontHandleSelIdxChg = True  ' Suppress selected index changed handler.
            InName = AppContext.BaseDirectory + DATA_DIR + "seartxtidx.dat"
            ' First line is word, second line is pointers to words in symptextdata.dat.
            Sfile_size = Val(System.IO.File.ReadAllLines(InName).Length.ToString())
            Dim SearchData() = System.IO.File.ReadAllLines(InName)
            GlobalSearchData = SearchData   ' Load file data to global variable for use by search functions
            For ptr1 = 0 To Sfile_size - 1 Step 2   ' This took 10 - 11 seconds
                Find.searData.Add(New searDat(SearchData(ptr1), SearchData(ptr1 + 1)))
            Next
            FileSystem.FileClose(Form1.SdxFile)
        Catch
            Dim unused3 = MsgBox(Prompt:="seartxtidx.dat file missing or corrupted; please re-install Dr. Hahnemann's Assistant to fix this problem",
            Buttons:=vbOKOnly + vbCritical,
            Title:="ERROR!")
            Close()
        End Try

        Try
            InName = AppContext.BaseDirectory + DATA_DIR + "onesymptextdata.dat"
            LineCount = Val(System.IO.File.ReadAllLines(InName).Length.ToString())
            Dim OneSympData() = System.IO.File.ReadAllLines(InName)
            For ptr1 = 0 To LineCount - 1 Step 2   ' This took 10 - 11 seconds
                OneSymp.oneSympData.Add(New oneSympDat(OneSympData(ptr1), OneSympData(ptr1 + 1)))
            Next
            FileSystem.FileClose(Form1.SdxFile)
        Catch
            Dim unused3 = MsgBox(Prompt:="onesymptextdata.dat file missing or corrupted; please re-install Dr. Hahnemann's Assistant to fix this problem",
                    Buttons:=vbOKOnly + vbCritical,
                    Title:="ERROR!")
            Close()
        End Try

        Try
            ' Pull in Thes data
            Form1.Syfile_size = Val(System.IO.File.ReadAllLines(AppContext.BaseDirectory + DATA_DIR + "thestxtidx.dat").Length.ToString())
            Dim loc_ThesData() = System.IO.File.ReadAllLines(AppContext.BaseDirectory + DATA_DIR + "thestxtidx.dat")
            ReDim QuestProgress.ThesWords(Form1.Syfile_size / 2)
            ReDim QuestProgress.ThesPointers(Form1.Syfile_size / 2)

            ptr2 = 0
            For ptr1 = 0 To Form1.Syfile_size - 1 Step 2 ' copy local data to public variable
                QuestProgress.ThesWords(ptr2) = loc_ThesData(ptr1)
                QuestProgress.ThesPointers(ptr2) = loc_ThesData(ptr1 + 1)
                ptr2 += 1
            Next
        Catch
            Dim unused3 = MsgBox(Prompt:="thestxtidx.dat file missing or corrupted; please re-install Dr. Hahnemann's Assistant to fix this problem",
            Buttons:=vbOKOnly + vbCritical,
            Title:="ERROR!")
            Close()
        End Try

        Try
            ErrorFlag = 1146
            ' Pull in constants
            SMaxSympPtr = SCombSympPtr

            DeleteMustButton.Enabled = False
            DeleteSelButton.Enabled = False
            PrescribeButton.Enabled = False
            Dirty = False       ' Initialize need-to-save-data flag.
            QDirty = False      ' Initialize need-to-save questionaire flag.

            If Not ExitPressed Then
                '   Initialize the symptom selections.
                MustBox.Text = ""
                MustData = ""
                ' Set up initial file names and display them.
                CaseFileName = "Untitled"
                QuestFileName = "Untitled"
                WriteTitles()     ' Display file names at top of screen.
            Else
                Close()
            End If

            ' Create a map of symptom ID's in the Lst1 indexes
            ReDim MapSymIDToIndex(26771)    ' The highest SymptomID is 26771; this is > the number of symptoms in Lst1
            For Ptr = 0 To Lst1.Items.Count - 1
                SymptomID = GetSymptomID(CombSympData(Ptr).SympData.ToString())
                MapSymIDToIndex(SymptomID) = Ptr
            Next

            ' Calling the .Show method for sub-forms here rather than in their associated button handler will improve performance
            ' on slower computers as the delay caused by loading their list-box databases is moved to program startup where it
            ' is less noticeable.  It also gets around an issue with list box .DisplayMember not re-initializing properly after
            ' the list box's sub-form is closed and re-opened.
            OneSymp.Show()
            OneSymp.Visible = False
            Find.Show()
            Find.Visible = False
            GridSearch.Show()
            GridSearch.Visible = False
            SelRem.Show()
            SelRem.Visible = False
            InLoad = False
        Catch
            Dim unused = MsgBox(Prompt:="Oops, something went wrong; program will exit. (" + Str(ErrorFlag) + ")",
            Buttons:=vbOKOnly + vbCritical,
            Title:="ERROR!")
            CloseAll()
            Close()
        End Try

        Try
            ' Pull in Remedy data
            Remfile_size = Val(System.IO.File.ReadAllLines(AppContext.BaseDirectory + DATA_DIR + "RemPresc.dat").Length.ToString())
            Dim loc_RemLines() = System.IO.File.ReadAllLines(AppContext.BaseDirectory + DATA_DIR + "RemPresc.dat")
            ReDim RemText(Form1.Remfile_size / 2)
            ReDim RemData(Form1.Remfile_size / 2)

            ptr2 = 0
            For ptr1 = 0 To Form1.Remfile_size - 1 Step 2 ' copy local data to public variable
                RemText(ptr2) = loc_RemLines(ptr1)
                RemData(ptr2) = loc_RemLines(ptr1 + 1)
                ptr2 += 1
            Next

        Catch
            Dim unused4 = MsgBox(Prompt:="RemPresc.dat file missing or corrupted; please re-install Dr. Hahnemann's Assistant to fix this problem",
            Buttons:=vbOKOnly + vbCritical,
            Title:="ERROR!")
            Close()
        End Try

    End Sub

    Private Sub Form1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If MouseButtons = 2 Then FileMenuItem.PerformClick()
    End Sub

    Public Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        ' Handles press of <CTRL> and any other key pressed in combination with it
        If e.KeyValue = CTRL_KEY Then
            CtrlPressed = True
        ElseIf e.KeyValue = ALT_KEY Then
            AltPressed = True
        Else

            If CtrlPressed Then ' Check for shortcut key press and dispatch appropriate function
                CtrlPressed = False ' In case of multiple KeyDown events without KeyUp, due to holding key down causing it to repeat
                Select Case e.KeyValue

                    Case Else
                        ' Default
                End Select

            ElseIf AltPressed Then
                AltPressed = False   ' In case of multiple KeyDown events without KeyUp, due to holding key down causing it to repeat
                Select Case e.KeyValue
                    Case B_KEY
                        Handle_CmdRemedies()

                    Case F_KEY
                        EditFind.Focus()
                        Handle_EditFind()

                    Case G_KEY
                        GridSearchButton.Focus()
                        Handle_GridSearchClick()

                    Case I_KEY
                        ShowSimilarSymptoms.Focus()
                        If ShowSimilarSymptoms.Enabled Then Handle_ShowSimilarSymptoms()

                    Case M_KEY
                        HandleMustUse()

                    Case Q_KEY
                        Command4.Focus()
                        Handle_Command4()

                    Case R_KEY
                        Rep.Focus()
                        Handle_Rep()

                    Case S_KEY
                        HandleSelRemedy()

                End Select
            End If
        End If
    End Sub

    Public Sub Form1_KeyUp(sender As Object, e As KeyEventArgs) Handles MyBase.KeyUp
        ' Check for release of <CTRL>; update CtrlPressed
        If e.KeyValue = CTRL_KEY Then
            CtrlPressed = False
        ElseIf e.KeyValue = ALT_KEY Then
            AltPressed = False
        End If
    End Sub
    Private Sub Form1_Unload(sender As Object, e As EventArgs) Handles MyBase.Closing
        If Not FileExitCalled Then FileExit()
    End Sub

    Public Sub TurnOnToolTips()
        ToolTip1.Active = True
        AutoWeightSeverity.ToolTip1.Active = True
        AutoWeightType.ToolTip1.Active = True
        Find.ToolTip1.Active = True
        FindAll.ToolTip1.Active = True
        GridSearch.ToolTip1.Active = True
        OneSymp.ToolTip1.Active = True
        PrescRem.ToolTip1.Active = True
        PrintSelections.ToolTip1.Active = True
        Questionnaire.ToolTip1.Active = True
        QuestProgress.ToolTip1.Active = True
        SelRem.ToolTip1.Active = True
        SetAutowt.ToolTip1.Active = True
        SetPref.ToolTip1.Active = True
    End Sub

    Public Sub TurnOffToolTips()
        ToolTip1.Active = False
        AutoWeightSeverity.ToolTip1.Active = False
        AutoWeightType.ToolTip1.Active = False
        Find.ToolTip1.Active = False
        FindAll.ToolTip1.Active = False
        GridSearch.ToolTip1.Active = False
        OneSymp.ToolTip1.Active = False
        PrescRem.ToolTip1.Active = False
        PrintSelections.ToolTip1.Active = False
        Questionnaire.ToolTip1.Active = False
        QuestProgress.ToolTip1.Active = False
        SelRem.ToolTip1.Active = False
        SetAutowt.ToolTip1.Active = False
        SetPref.ToolTip1.Active = False
    End Sub

    Public Function GetSymptomID(dataString As String) As Integer
        ' Returns symptom ID, or 0 if no valid ID found, contained in data string; 1st number is symptom ID
        Dim StrPtr As Integer   ' string pointer
        Dim IDString As String  ' string of ID's
        Dim IntID As Integer

        StrPtr = InStr(dataString, ",") 'Find end of ID's string.
        If (StrPtr <> 0) Then
            IDString = Strings.Left(dataString, StrPtr - 1)
            IntID = Val(IDString)   ' Convert it to integer.
        Else
            IntID = 0   ' Cover case where string doesn't contain an ID (should never happen)
        End If
        GetSymptomID = IntID
    End Function

    Public Function GetSymptomWeight(sympString As String) As Integer
        ' Returns the weight from a selected symptom (1 - 6) or 0 if no valid weight in string
        Dim WeightString As String  ' String from of weight
        Dim IntWeight As Integer    ' Integer form of weight

        WeightString = Strings.Left(sympString, 1)
        If (WeightString = "1" Or WeightString = "2" Or WeightString = "3" Or WeightString = "4" Or WeightString = "5" Or WeightString = "6") Then
            IntWeight = Val(WeightString)   ' Convert it to integer.
        Else
            IntWeight = 0   ' Cover case where string doesn't contain a valid weight (should never happen)
        End If

        GetSymptomWeight = IntWeight
    End Function

    Public Function IsSymptomSelected(sympIndex As Integer) As Boolean
        ' Returns True if symptom has been selected in the Lst1 (symptom selection) box, otherwise False.
        Dim ptr As Integer

        Try
            RaiseException = False
            IsSymptomSelected = False
            For ptr = 0 To SymSelectedSize - 1
                If sympIndex = SymSelected(ptr) Then
                    IsSymptomSelected = True
                    Exit For
                End If
            Next

        Catch
            IsSymptomSelected = False
            RaiseException = True
        End Try
    End Function

    Public Function IsSymptomInSelectedList(strSelLstDat As String) As Boolean
        ' Returns True if symptom has been selected in the Selected List box, otherwise False.
        Dim ptr As Integer
        Dim compareData As String

        Try
            RaiseException = False
            IsSymptomInSelectedList = False
            '        If SelLst.Items.Count > 0 And SelLstData(0).SelSympDesc.ToString <> " " Then
            If SelLst.Items.Count > 0 Then
                If SelSympData(0).SelSympDesc.ToString <> " " Then
                    For ptr = 0 To SelLst.Items.Count - 1
                        compareData = SelSympData(ptr).SelSympData.ToString
                        If strSelLstDat = compareData Then
                            IsSymptomInSelectedList = True
                            Exit For
                        End If
                    Next
                End If
            End If

        Catch
            IsSymptomInSelectedList = False
            RaiseException = True
        End Try
    End Function

    Public Function IsSymptomInMustUseList(sympID As Integer) As Boolean
        ' Returns True if symptom has been selected in the Must Use box, otherwise False.
        Dim ptr As Integer

        Try
            RaiseException = False
            IsSymptomInMustUseList = False
            For ptr = 0 To MustListSize - 1
                If sympID = MustList(ptr) Then
                    IsSymptomInMustUseList = True
                    Exit For
                End If
            Next

        Catch
            IsSymptomInMustUseList = False
            RaiseException = True
        End Try
    End Function

    Public Function AddSymptomToMustUseList(sympID As Integer) As Boolean
        ' Adds tymptom ID to MustList array if not already in it.
        Try
            If Not (IsSymptomInSelectedList(sympID)) Then
                MustListSize += 1
                ReDim Preserve MustList(MustListSize)
                MustList(MustListSize - 1) = sympID
                AddSymptomToMustUseList = True
            Else
                AddSymptomToMustUseList = False
            End If

        Catch
            AddSymptomToMustUseList = False
            Form1.RaiseException = True
        End Try
    End Function

    Public Function RemoveSymptomFromMustUseList(sympID As Integer) As Boolean
        ' Looks for sympID in MustUseList array, sets it to 0 and returns True if found, otherwise returns False.
        Dim ptr As Integer
        RemoveSymptomFromMustUseList = False
        For ptr = 0 To MustListSize - 1
            If sympID = MustList(ptr) Then
                MustList(ptr) = 0
                RemoveSymptomFromMustUseList = True
                Exit For
            End If
        Next
    End Function

    Private Sub NewCaseFile_Click(sender As Object, e As EventArgs) Handles NewCaseFile.Click
        FileNew()
    End Sub

    Private Sub OpenCaseFile_Click(sender As Object, e As EventArgs) Handles OpenCaseFile.Click
        FileOpen()
    End Sub

    Private Sub SaveCaseFile_Click(sender As Object, e As EventArgs) Handles SaveCaseFile.Click
        If CaseFileName = "Untitled" Then
            FileSaveAs()
        Else
            SaveCurrentFile()
        End If
    End Sub

    Private Sub SaveCaseFileAs_Click(sender As Object, e As EventArgs) Handles SaveCaseFileAs.Click
        FileSaveAs()
    End Sub

    Private Sub CloseCaseFile_Click(sender As Object, e As EventArgs) Handles CloseCaseFile.Click
        FileClose()
    End Sub

    Private Sub NewQuestFile_Click(sender As Object, e As EventArgs) Handles NewQuestFile.Click
        FileQuestNew()
    End Sub

    Private Sub OpenQuestFile_Click(sender As Object, e As EventArgs) Handles OpenQuestFile.Click
        FileQuestOpen()
        WriteQ1Screen()
        WriteQ2Screen()
        WriteQ3Screen()
        WriteQ4Screen()
        WriteQ5Screen()
        WriteQ6Screen()
        WriteQ7Screen()
        Form1.QDirty = False
    End Sub

    Private Sub SaveQuestFile_Click(sender As Object, e As EventArgs) Handles SaveQuestFile.Click
        If QuestFileName = "Untitled" Then
            FileQuestSaveAs()
        Else
            SaveCurrentQuestFile()
        End If
    End Sub

    Private Sub SaveQuestFileAs_Click(sender As Object, e As EventArgs) Handles SaveQuestFileAs.Click
        FileQuestSaveAs()
    End Sub

    Private Sub CloseQuestFile_Click(sender As Object, e As EventArgs) Handles CloseQuestFile.Click
        FileQuestClose()
    End Sub

    Private Sub Print_Click(sender As Object, e As EventArgs) Handles Print.Click
        Dim NumSelections As Integer    ' Number of enabled selections, used to determine if need to show PrintSelections
        Dim pdResult As DialogResult

        stringBuffer = ""   ' Clear out print buffer
        PrintPreviewShow = False
        NumSelections = 0
        If SelLst.Items.Count > 0 Then NumSelections += 1
        If PrescRem.Visible = True Then NumSelections += 2
        If Questionnaire.TextBox1.Text <> "" Or Questionnaire.TextBox2.Text <> "" Then NumSelections += 1
        If NumSelections > 1 Then
            PrintSelections.Show()  ' More than one possible print item, prompt for item(s) to print and call print handlers.
        ElseIf NumSelections = 1 And (Questionnaire.TextBox1.Text <> "" Or Questionnaire.TextBox2.Text <> "") Then
            QPrint = True
            PrintSelections.HandlePrint()
        Else    ' Just need to print selected symptoms list
            FilePrint() ' Load print buffer

            PrintDialog1.AllowSomePages = True   ' Allow the user to choose the page range he or she would like to print.
            PrintDialog1.AllowCurrentPage = True
            PrintDialog1.AllowPrintToFile = True
            PrintDialog1.AllowSelection = True
            PrintDialog1.ShowHelp = True     ' Show the help button.

            pdResult = PrintDialog1.ShowDialog()
            If pdResult = vbOK Then
                SetupAndPrint()
                PrintDocument1.Print()
            End If
            PrintDialog1.Dispose()
        End If

    End Sub

    Private Sub PrintPreview_Click(sender As Object, e As EventArgs) Handles PrintPreview.Click
        Dim NumSelections As Integer    ' Number of enabled selections, used to determine if need to show PrintSelections

        stringBuffer = ""   ' Clear out print buffer
        PrintPreviewShow = True
        NumSelections = 0
        If SelLst.Items.Count > 0 Then NumSelections += 1
        If PrescRem.Visible = True Then NumSelections += 2
        If Questionnaire.TextBox1.Text <> "" Or Questionnaire.TextBox2.Text <> "" Then NumSelections += 1
        If NumSelections > 1 Then
            PrintSelections.Show()  ' More than one possible print item, prompt for item(s) to print and call print handlers.
        Else    ' Just need to print selected symptoms list
            FilePrint() ' Load print buffer

            ' Get printer font from registry
            PrinterFontName = New Font(GetSetting("Hahnasst", "Startup", "PrinterFontName", [Default]:="Microsoft Sans Serif"),
            GetSetting("Hahnasst", "Startup", "PrinterFontSize", [Default]:="8"))
            PrintPreviewDialog1.ShowIcon = False
            PrintPreviewDialog1.WindowState = FormWindowState.Normal
            PrintPreviewDialog1.StartPosition = FormStartPosition.CenterScreen
            PrintPreviewDialog1.ClientSize = New Size(800, 1200)
            PrintPreviewDialog1.UseAntiAlias = True
            PrintPreviewDialog1.Document = PrintDocument1
            PrintPreviewDialog1.ShowDialog()
            PrintPreviewDialog1.Close()
            PrintPreviewDialog1.Dispose()
        End If
    End Sub

    Private Sub ExitApp_Click(sender As Object, e As EventArgs) Handles ExitApp.Click
        Close()
    End Sub


    Private Sub Normalize_Click(sender As Object, e As EventArgs)
        OptionsNormalize()
    End Sub



    Private Sub AutomaticallyAssignWeight_Click(sender As Object, e As EventArgs)
        OptionsAutoWeight()
    End Sub


    Private Sub HelpContents_Click(sender As Object, e As EventArgs) Handles HelpMenuContents.Click
        Dim msg As String
        Dim myProcess As New Process

        Try
            myProcess.StartInfo.FileName = AppContext.BaseDirectory + HELP_DIR + "HAHNASST.chm"
            myProcess.StartInfo.UseShellExecute = True
            myProcess.StartInfo.RedirectStandardOutput = False
            myProcess.Start()
            myProcess.Dispose()
        Catch
            msg = "Help file not found.  Please re-install program to fix this problem."
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub

    Private Sub AboutHahnasst_Click(sender As Object, e As EventArgs) Handles AboutHahnasst.Click
        About.Show()
    End Sub

    Private Sub ToolStripFileNew_Click(sender As Object, e As EventArgs) Handles ToolStripFileNew.Click
        If ToolStripFileNew.Enabled = True Then
            FileNew()
        End If
    End Sub

    Private Sub ToolStripFileOpen_Click(sender As Object, e As EventArgs) Handles ToolStripFileOpen.Click
        If ToolStripFileOpen.Enabled = True Then
            FileOpen()
        End If
    End Sub

    Private Sub ToolStripFileSave_Click(sender As Object, e As EventArgs) Handles ToolStripFileSave.Click
        If ToolStripFileSave.Enabled = True Then
            SaveCurrentFile()
        End If
    End Sub

    Private Sub ToolStripFileSaveAs_Click(sender As Object, e As EventArgs) Handles ToolStripFileSaveAs.Click
        If ToolStripFileSaveAs.Enabled = True Then
            FileSaveAs()
        End If
    End Sub

    Private Sub ToolStripFileClose_Click(sender As Object, e As EventArgs) Handles ToolStripFileClose.Click
        If ToolStripFileClose.Enabled = True Then
            FileClose()
        End If
    End Sub

    Private Sub ToolStripExit_Click(sender As Object, e As EventArgs) Handles ToolStripExit.Click
        FileExit()
    End Sub

    Private Sub ToolStripPrint_Click(sender As Object, e As EventArgs) Handles ToolStripPrint.Click
        If ToolStripPrint.Enabled = True Then
            FilePrint()
        End If
    End Sub

    Private Sub ToolStripHelp_Click(sender As Object, e As EventArgs) Handles ToolStripHelp.Click
        Dim msg As String
        Dim myProcess As New Process

        Try
            myProcess.StartInfo.FileName = AppContext.BaseDirectory + HELP_DIR + "HAHNASST.chm"
            myProcess.StartInfo.UseShellExecute = True
            myProcess.StartInfo.RedirectStandardOutput = False
            myProcess.Start()
            myProcess.Dispose()
        Catch
            msg = "Help file not found.  Please re-install program to fix this problem."
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text = "" Then
            ShowSimilarSymptoms.Enabled = False
            SelectRemedy.Enabled = False
            MustUse.Enabled = False
        Else
            SelectRemedy.Enabled = True
            MustUse.Enabled = True
            ShowSimilarSymptoms.Enabled = True
        End If

    End Sub

    Private Sub Command4_Click(sender As Object, e As EventArgs) Handles Command4.Click ' Questionnaire
        Handle_Command4()
    End Sub

    Private Sub Handle_Command4()   ' Questionnaire
        Dim msg As String           ' error message string

        Try
            ErrorFlag = 1130
            If SympSize = 0 Then
                SympSize = 1
            End If
            SymPtr = 0
            Questionnaire.Top = 0
            Questionnaire.Left = 0
            Questionnaire.Show()

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub

    Private Sub Rep_Click(sender As Object, e As EventArgs) Handles Rep.Click
        Handle_Rep()
    End Sub

    Private Sub Handle_Rep()
        Dim msg As String           ' error message string

        Try
            ErrorFlag = 1131
            OneSymp.Visible = True

        Catch
            Cursor = Cursors.Default    ' Change cursor to default.
            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
               Buttons:=vbOKOnly + vbCritical,
               Title:="ERROR!")
        End Try
    End Sub

    Private Sub Lst1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Lst1.SelectedIndexChanged
        Dim done As Boolean
        Dim found As Boolean    ' indicates missing item was found
        Dim idxPtr As Integer  ' index pointer
        Dim msg As String   ' error message string
        Dim selPtr As Integer   ' Lst1 selectedIndices pointer

        ErrorFlag = 1132
        Try
            If Not InLoad Then
                FirstLoad = False
                ' Need to detect if index was selected or de-selected:
                'if selected, add to end of symSelected array
                'if de-selected, remove entry from symSelected array and move subsequent items up
                'NOTE:  This code needs to be here rather than in "_Click" to handle setting of
                'Lst1 selected from other places in the code.
                If Lst1.SelectedIndices.Count > SymSelectedSize Then    ' A new item was selected
                    ' Find the new item by comparing each Lst1.SelectedIndices against the SymSelected array
                    done = False
                    idxPtr = 0
                    selPtr = 0
                    ' For each Lst1 selected index, look for a corresponding entry in SymSelected.
                    While Not done And selPtr < Lst1.SelectedItems.Count
                        idxPtr = 0
                        found = False
                        While Not found And idxPtr < SymSelectedSize
                            If Lst1.SelectedIndices(selPtr).Equals(SymSelected(idxPtr)) Then found = True
                            idxPtr += 1
                        End While
                        selPtr += 1
                        If Not found Then
                            done = True
                            ' Add its index to the end of SymSelected array
                            SymSelected(SymSelectedSize) = Lst1.SelectedIndices(selPtr - 1)
                            SymSelectedSize += 1
                            TextBox1.Text = CombSympData(Lst1.SelectedIndices(selPtr - 1)).SympDesc.ToString()

                        End If
                    End While
                End If
            End If
        Catch
            Cursor = Cursors.Default    ' Change cursor to default.
            msg = "Oops, something went wrong; please restart the program. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
               Buttons:=vbOKOnly + vbCritical,
               Title:="ERROR!")
        End Try
    End Sub
    Private Sub Lst1_Click(sender As Object, e As EventArgs) Handles Lst1.Click
        Dim found As Boolean    ' indicates missing item was found
        Dim idxPtr As Integer  ' SymSelected() index pointer
        Dim loopPtr As Integer  ' pointer used to move SelectedIndices elememts up
        Dim msg As String   ' error message string
        Dim selPtr As Integer   ' Lst1 selectedIndices pointer
        Dim symAlreadySelected As Boolean   ' Indicates symptom is in MustBox or SelLst
        Dim SymDatText As String    ' text of symptom data

        Try
            ErrorFlag = 1133

            If Lst1.SelectedIndices.Count < SymSelectedSize Then    ' An item was de-selected
                ' Find the de-selected item by comparing SymSelected array against Lst1.SelectedIndices
                ' Find the new item by comparing each Lst1.SelectedIndices against the SymSelected array
                'NOTE:  This code needs to be here rather than in "_SelectedIndexChanged" to prevent
                ' recursion when an item is re-selected from the code.
                found = True
                idxPtr = 0
                selPtr = 0
                ' Look for any de-selected symptom; for each Lst1 selected index, look for a corresponding entry in SymSelected.
                While found And idxPtr < SymSelectedSize
                    found = False
                    selPtr = 0
                    If Lst1.SelectedItems.Count > 0 Then
                        While Not found And selPtr < Lst1.SelectedItems.Count
                            If Lst1.SelectedIndices(selPtr).Equals(SymSelected(idxPtr)) Then found = True
                            selPtr += 1
                        End While
                        'Else    ' There will be only a single entry in SymSelected, which points to the de-selected Lst1 item.
                        ' idxPtr is already 0, which points to the entry
                        ' found should remain false
                    End If
                    idxPtr += 1
                End While

                If Not found Then   ' selPtr will be one plus the index of the "not found" (deselected) item in SymSelected() array
                    idxPtr -= 1
                    SymDatText = Lst1.Items((SymSelected(idxPtr))).SympData.ToString
                    ' If un-selected symptom is in MustBox or SelLst, show message and re-select it.
                    symAlreadySelected = False
                    If MustBox.Text <> "" And SymDatText = MustData Then
                        symAlreadySelected = True
                        Dim unused = MsgBox(Prompt:="Symptom is in 'Must Use' box; use 'Delete Must' button to un-select it.",
                        Buttons:=vbOKOnly + vbInformation,
                        Title:="INFORMATION")
                        Lst1.SetSelected(SymSelected(idxPtr), True)
                    Else
                        If SelLst.Items.Count > 0 Then
                            ' SymSelected() is the array of all symptoms in Lst1 that are selected in Lst1
                            If IsSymptomInSelectedList(SymDatText) Then
                                symAlreadySelected = True
                                Dim unused = MsgBox(Prompt:="Symptom is in 'SELECTED SYMPTOMS' list; use 'Delete Selected' button to un-select it.",
                                      Buttons:=vbOKOnly + vbInformation,
                                      Title:="INFORMATION")
                                Lst1.SetSelected(SymSelected(idxPtr), True)
                            End If
                        End If
                        If symAlreadySelected = False Then  ' Symptom is not in Select List
                            ' Remove its index from the SymSelected array; move remaining items up in the array
                            If selPtr > 0 Then selPtr -= 1
                            For loopPtr = idxPtr To SymSelectedSize - 1
                                SymSelected(loopPtr) = SymSelected(loopPtr + 1)
                            Next
                            ' Remove last element from SymSelected() array
                            SymSelected(SymSelectedSize - 1) = -1
                            SymSelectedSize -= 1
                            If SymSelectedSize > 0 Then
                                TextBox1.Text = CombSympData(Lst1.SelectedIndices(SymSelectedSize - 1)).SympDesc.ToString()
                            Else
                                TextBox1.Text = ""
                            End If
                        End If
                    End If
                End If
            End If
            If Lst1.SelectedItems.Count = 0 Then  ' If no items are selected
                DisableFilePrintPreview()
                DisableFilePrint()
                SelectRemedy.Enabled = False    ' Disable Select Remedy Button
                MustUse.Enabled = False         ' Disable Must Use button
                ShowSimilarSymptoms.Enabled = False ' Disable Show Similar Symptoms
            Else
                SelectRemedy.Enabled = True    ' Enable Select Remedy Button
                MustUse.Enabled = True         ' Enable Must Use button
                ShowSimilarSymptoms.Enabled = True ' Enable Show Similar Symptoms
            End If
        Catch
            Cursor = Cursors.Default    ' Change cursor to default.
            msg = "Oops, something went wrong; please restart the program. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub

    Public Function InSelectedIndicesArray(SelIndex As Integer) As Boolean
        ' If SelIndex is contained in the SymSelected array it returns True, otherwise False
        InSelectedIndicesArray = False
        For listItem = 0 To SymSelectedSize
            If SymSelected(listItem) = SelIndex Then
                InSelectedIndicesArray = True
            End If
        Next
    End Function


    ' Action:  Keep selections in Lst1 unless clicked on again.  Clicking on Select adds all selections not in Selected Symptoms
    ' or Must Use to Selected Symptoms.  Similar for clicking Must Use.
    ' Text box always shows last symptom selected in Lst1.  If that symptom is un-selected it shows first selected symptom.
    ' Selected Symptoms and Must Use are single selection only.
    ' Deleting from Selected or Must also un-selects from Lst1.
    ' If one or more symptoms in Must Use, selecting another for Must Use brings up dialog "Replace current must-use selection?"
    ' No adds it to the Must Use; Yes should replace selected Must symptom, but has a bug.  DON'T USE; allow more than 
    ' one item in Must Use list.
    Private Sub DeleteMustButton_Click(sender As Object, e As EventArgs) Handles DeleteMustButton.Click     ' Delete Must
        Dim IntID As Integer        ' integer form of symptom ID
        Dim IDFound As Boolean      ' item to delete was found
        Dim msg As String           ' error message string
        Dim selPtr As Integer       ' Lst1 selected index pointer

        Try
            ErrorFlag = 1033

            IntID = GetSymptomID(MustBox.Text)   ' Get symptom ID of MustBox.
            IDFound = False

            msg = "Move current must-use selection to selected list?"
            If (MsgBox(msg, vbYesNo + vbQuestion) = 6) Then 'Yes was clicked
                ErrorFlag = 1034
                ' Move item to SelLst; no need to add weight because it's already in the string.
                SelSympData.Add(New SelSympDat(MustBox.Text, MustData))
                If NoSelectedItems Then ' If this is the first item in SelLst
                    ' Set up data source for SelLst
                    SelLst.DataSource = SelSympData
                    SelLst.ValueMember = "SelSympData"
                    SelLst.DisplayMember = "SelSympDesc"
                    SelLst.ClearSelected()
                    NoSelectedItems = True
                    SelLst_UpdateDataSource()   ' Make it appear in the list box
                Else
                    SortAndUpdateSelLst()   ' Sort the Arraylist by symptom text and update SelLst box
                End If

                ' Remove selected item from MustBox.
                MustBox.Text = ""
                MustData = ""
                DeleteSelButton.Enabled = True

            Else    ' Deselect symptom in Lst1

                ErrorFlag = 1035
                IntID = GetSymptomID(MustData)   ' Get symptom ID of "Must" item.

                ' Remove selected item from MustBox.
                MustBox.Text = ""
                MustData = ""

                ErrorFlag = 1036
                For selPtr = 0 To Lst1.SelectedItems.Count - 1
                    If GetSymptomID(CombSympData(Lst1.SelectedIndices(selPtr)).SympData.ToString) = IntID Then
                        Lst1.SetSelected(Lst1.SelectedIndices(selPtr), False) 'Deselect the string just deleted.
                        Lst1_Click(sender, e)
                        Exit For
                    End If
                Next selPtr

                If (SelLst.Items.Count = 0) And (MustBox.Text = "") Then
                    DisableFilePrintPreview()
                    DisableFilePrint()
                    If Lst1.SelectedIndex = -1 Then ' Item was de-selected
                        If Lst1.SelectedIndex - 1 Then  ' If no items are selected
                            SelectRemedy.Enabled = False    ' Disable Select Remedy Button
                            MustUse.Enabled = False         ' Disable Must Use button
                            ShowSimilarSymptoms.Enabled = False ' Disable Show Similar Symptoms Button
                        End If
                    End If
                End If
                Dirty = True

            End If

            DeleteMustButton.Enabled = False
            If (SelLst.Items.Count = 0) And (MustBox.Text = "") Then
                DisableFilePrintPreview()
                DisableFilePrint()
                PrescribeButton.Enabled = False
            End If
            Dirty = True

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub

    Private Sub CmdRemedies_Click(sender As Object, e As EventArgs) Handles CmdRemedies.Click
        Handle_CmdRemedies()
    End Sub

    Public Sub Handle_CmdRemedies()
        Dim msg As String           ' error message string

        Try
            SelRem.Top = 0
            SelRem.Left = 0
            SelRem.Visible = True

        Catch
            Cursor = Cursors.Default    ' Change cursor to default.

            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub

    Public Sub SelLst_UpdateDataSource()
        SelLst.DataSource = Nothing ' Remove and re-associate database in order to refresh SelLst
        SelLst.DataSource = SelSympData
        SelLst.ValueMember = "SelSympData"
        SelLst.DisplayMember = "SelSympDesc"
        SelLst.ClearSelected()
    End Sub


    Private Sub SelectRemedy_Click(sender As Object, e As EventArgs) Handles SelectRemedy.Click ' Select Symptom button
        HandleSelRemedy()
    End Sub

    Public Sub HandleSelRemedy()

        Dim IDString As String          ' string of ID's
        Dim msg As String               ' for diagnostics
        Dim Ptr As Integer              ' loop pointer

        Try
            ErrorFlag = 1037
            Form1.RaiseException = False
            If Lst1.SelectedIndex <> -1 Then
                For Ptr = 0 To SymSelectedSize - 1
                    ' Get Symptom ID
                    IDString = CombSympData(Lst1.SelectedIndices(Ptr)).SympData.ToString
                    HandleSelectRemedy(IDString)
                Next Ptr
            End If
        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try

    End Sub

    Public Sub HandleSelectRemedy(SymDatText As String)
        Dim IntID As Integer            ' integer form of symptom ID
        Dim locRemID As Integer         ' local remedy ID
        Dim locSympDesc As String       ' local copy of symptom description string
        Dim StrPtr As Integer           ' string pointer
        Dim StrRemID As String          ' string form of Remedy ID
        Dim sw As Integer               ' symptom weight value

        ' If symptom is not already selected, set it selected in Form1.Lst1, add symptom weight to its description text and add this and SymDatText to SelLst.
        Form1.ErrorFlag = 4003
        IntID = Val(SymDatText)
        ' Get Remedy ID
        StrPtr = InStr(SymDatText, ",")
        StrRemID = SymDatText.Substring(StrPtr)
        locRemID = Val(StrRemID)

        ' Check that symptom not already selected.
        If (Not IsSymptomInSelectedList(SymDatText) And Not (SymDatText = Form1.MustData)) Then
            sw = 0
            locSympDesc = Lst1.Items(MapSymIDToIndex(IntID)).SympDesc.ToString
            If Form1.AutoWeight Then
                AutoWeightText = locSympDesc    ' Pass string to AutoWeightSeverity form
                AutoWeightTp = GetAutoweight(locSympDesc) ' Get symptom-specific weighting factor:  1 for physical symptoms, 2 for general symptoms, or 3 for mind symptoms, 0 if unable to classify
                If AutoWeightTp = 0 Then  ' need to manually assign
                    AutoWeightType.ShowDialog() ' Get AutoWeightTp from dialog
                End If
                AutoWeightSeverity.ShowDialog() ' Get AutoWeightSev from dialog
                sw = AutoWeightTp + AutoWeightSev   ' Symptom Weight is sum of symptom type factor plus symptom severity factor
                locSympDesc = Str$(sw) + Chr(9) + locSympDesc
                locSympDesc = Mid(locSympDesc, 2)
            Else
                ' Use weight from "weight" combo box
                locSympDesc = Weight.Text + Chr(9) + locSympDesc
                sw = Weight.Text
            End If
            ErrorFlag = 4004
            If InStr(locSympDesc, "  ") <> 0 Then ' If locSympDesc is from questionaire.
                locSympDesc = Mid(locSympDesc, 1, 1) + Mid(locSympDesc, InStr(locSympDesc, "  ") + 2)
                locSympDesc = Mid(locSympDesc, 1, 1) + Chr(9) + Mid(locSympDesc, 2)
            End If

            SelSympData.Add(New SelSympDat(locSympDesc, SymDatText))

            If NoSelectedItems Then ' If this is the first item in SelLst
                ' Set up data source for SelLst
                SelLst.DataSource = Form1.SelSympData
                SelLst.ValueMember = "SelSympData"
                SelLst.DisplayMember = "SelSympDesc"
                SelLst.ClearSelected()
                NoSelectedItems = False
                SelLst_UpdateDataSource()   ' Make it appear in the list box
            Else
                SortAndUpdateSelLst()   ' Sort the Arraylist by symptom text and update SelLst box
            End If
            Lst1.SetSelected(Form1.MapSymIDToIndex(IntID), True)   ' Select it in Form1 Symptoms list box

            If (UpdateRemList(locRemID, False)) Then
                DeleteSelButton.Enabled = True
                PrescribeButton.Enabled = True
            End If
        End If
        If Form1.RaiseException Then Throw New System.Exception("An exception has occurred.")
        EnableFileSave()
        EnableFileSaveAs()
        EnableFilePrint()
        Form1.Dirty = True

    End Sub

    Public Function GetLst1SelectedItem(SymptomID As Integer) As String
        ' Returns the Lst1 item that has SymptomID.
        Dim Ptr As Integer
        Dim locID As Integer

        GetLst1SelectedItem = "NULL"    ' Handle case where item is not found; should never happen.
        For Ptr = 0 To Lst1.SelectedItems.Count - 1
            locID = GetSymptomID(Lst1.SelectedItems(Ptr))
            If locID = SymptomID Then
                GetLst1SelectedItem = Lst1.SelectedItems(Ptr)
                Exit Function
            End If
        Next
    End Function

    Private Sub MustUse_Click(sender As Object, e As EventArgs) Handles MustUse.Click     ' Must Use button
        HandleMustUse()
    End Sub

    Public Sub HandleMustUse()

        Dim IDString As String          ' string of ID's
        Dim msg As String               ' for diagnostics
        Dim SymptomText As String       ' Text of the symptom

        Try
            ErrorFlag = 1039
            Form1.RaiseException = False
            If Lst1.SelectedIndex <> -1 Then
                SymptomText = TextBox1.Text
                IDString = Lst1.Items(SymSelected(SymSelectedSize - 1)).SympData.ToString
                Handle_MustUse(IDString, SymptomText)
            End If
            Exit Sub

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub

    Public Sub Handle_MustUse(SympIDString As String, SympText As String)
        Dim ComPtr As Integer           ' comma pointer
        Dim IntID As Integer            ' integer form of symptom ID
        Dim locSympDesc As String       ' local copy of symptom description string
        Dim msg As String               ' for diagnostics
        Dim MustDat As String           ' data read from Must Box
        Dim OldIntID As Integer         ' integer form of old "Must Use" symptom ID from MustData
        Dim SelPtr As Integer           ' pointer to selected symptom array
        Dim SW As Integer               ' symptom weight value
        Dim Temp As String              ' temporary string
        Dim UserWeight As Integer       ' user-selected symptom weight, 1-6

        locSympDesc = " "   ' Avoid warning variable used before assigning value

        ' Check that symptom not already selected.
        If (Not IsSymptomInSelectedList(SympIDString) And Not SympIDString = MustData) Then
            DeleteMustButton.Enabled = True
            PrescribeButton.Enabled = True
            IntID = Val(SympIDString)
            Temp = Lst1.Items(MapSymIDToIndex(IntID)).SympDesc.ToString
            ' If symptom is from questionnaire, strip off page and question numbers.
            If Mid(Temp, 2, 1) = "/" Then
                ComPtr = InStr(Temp, ",")
                Temp = Mid(Temp, ComPtr + 1)
                ComPtr = InStr(Temp, ",")
                Temp = Mid(locSympDesc, ComPtr + 3)
            End If

            locSympDesc = Temp

            If AutoWeight Then
                AutoWeightText = Temp    ' Pass string to AutoWeightSeverity form

                SW = 0
                AutoWeightTp = GetAutoweight(Temp)   ' Get symptom-specific weighting factor:  1 for physical symptoms, 2 for general symptoms, or 3 for mind symptoms, 0 if unable to classify
                If AutoWeightTp = 0 Then  ' need to manually assign
                    AutoWeightType.ShowDialog() ' Get AutoWeightTp from dialog
                End If
                AutoWeightSeverity.ShowDialog()
                SW = AutoWeightTp + AutoWeightSev
            Else
                ' Use weight from "weight" combo box
                locSympDesc = Weight.Text + Chr(9) + locSympDesc
                SW = Weight.Text
            End If

            If MustBox.Text <> "" Then   ' If the MustBox already contains a string.
                msg = "Replace current must-use selection?"
                ' Note if "No" is selected, nothing will be done; selection will remain selected in Lst1.
                If (MsgBox(msg, vbYesNo + vbQuestion) = 6) Then ' If "Yes" was selected
                    msg = "Move current must-use selection to selected list?"
                    If (MsgBox(msg, vbYesNo + vbQuestion) = 6) Then ' If "Yes" was selected
                        ErrorFlag = 1040
                        MustDat = MustBox.Text  ' Save off symptom text for moving to Selected List
                        UserWeight = Int(Strings.Left(MustDat, 1))
                        IntID = GetSymptomID(MustData)
                        If (IntID <> 0) Then
                            ' AddSymptomToSelectList(SelIntID)
                            ErrorFlag = 1041
                            MustBox.Text = ""
                            ErrorFlag = 1042
                            SelSympData.Add(New SelSympDat(MustDat, MustData))
                            If NoSelectedItems Then ' If this is the first item in SelLst
                                ' Set up data source for SelLst
                                SelLst.DataSource = SelSympData
                                SelLst.ValueMember = "SelSympData"
                                SelLst.DisplayMember = "SelSympDesc"
                                SelLst.ClearSelected()
                                NoSelectedItems = False
                            End If
                            SelLst_UpdateDataSource()   ' Make it appear in the list box
                            Lst1.SetSelected(Form1.MapSymIDToIndex(IntID), True)   ' Select it in Form1 Symptoms list box
                            DeleteSelButton.Enabled = True
                            Dirty = True
                        End If
                        MustBox.Text = locSympDesc
                        MustData = LastSelectedDataString

                    Else    ' No, don't move to selected list; need to unselect its Lst1 entry.
                        ErrorFlag = 1118
                        MustBox.Text = locSympDesc
                        OldIntID = GetSymptomID(MustData)  ' Old "Must Use" symptom ID, for de-selecting in Lst1
                        MustData = LastSelectedDataString
                        For SelPtr = 0 To Lst1.SelectedItems.Count - 1
                            If GetSymptomID(CombSympData(Lst1.SelectedIndices(SelPtr)).SympData.ToString) = OldIntID Then
                                Lst1.SetSelected(Lst1.SelectedIndices(SelPtr), False) 'Deselect the string just deleted.
                                Exit For
                            End If
                        Next SelPtr
                        Dirty = True
                    End If
                End If
            Else    ' Nothing in MustBox, OK to add
                If SymSelected(SymSelectedSize - 1) <> -1 Then ' don't add item if it has been de-selected
                    MustBox.Text = locSympDesc
                    MustData = Lst1.Items(SymSelected(SymSelectedSize - 1)).SympData.ToString
                    IntID = Val(MustData)
                    Lst1.SetSelected(Form1.MapSymIDToIndex(IntID), True)   ' Select it in Form1 Symptoms list box
                    Dirty = True
                End If
            End If
            EnableFilePrintPreview()
            EnableFilePrint()
            PrescribeButton.Enabled = True
        End If
    End Sub
    Private Sub EditFind_Click(sender As Object, e As EventArgs) Handles EditFind.Click

        Handle_EditFind()
    End Sub

    Private Sub Handle_EditFind()

        Dim msg As String           ' error message string

        Try
            ErrorFlag = 1131

            Find.Text1.Text = SearchString
            If Find.Text1.Text <> "" Then
                Find.FindNext.Enabled = True
                Find.FindFirst.Enabled = True
                Find.FindPrev.Enabled = True
                Find.FindAllButton.Enabled = True
                Find.FindNext.Select()
            Else
                Find.FindNext.Enabled = False
                Find.FindFirst.Enabled = False
                Find.FindPrev.Enabled = False
                Find.FindAllButton.Enabled = False
                Find.FindFirst.Select()
            End If

            Find.Visible = True
            Find.Text1.Select()
            Find.List1.Visible = False
        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub


    Private Sub GridSearchButton_Click(sender As Object, e As EventArgs) Handles GridSearchButton.Click

        Handle_GridSearchClick()
    End Sub

    Private Sub Handle_GridSearchClick()

        Dim msg As String   ' error message string
        Dim NeedEnabling As Boolean
        Dim Ptr As Integer  ' loop counter

        Try
            ErrorFlag = 1043
            NeedEnabling = False

            ' Restore last-used words.
            For Ptr = 0 To 23
                If GridString(Ptr) <> "" Then
                    GridSearch.Text1(Ptr).Text = GridString(Ptr)
                    NeedEnabling = True
                End If
            Next Ptr

            If NeedEnabling Then
                GridSearch.FindNext.Enabled = True
                GridSearch.FindFirst.Enabled = True
                GridSearch.FindPrev.Enabled = True
                GridSearch.FindAllButton.Enabled = True
                GridSearch.FindNext.Select()
            Else
                GridSearch.FindNext.Enabled = False
                GridSearch.FindFirst.Enabled = False
                GridSearch.FindPrev.Enabled = False
                GridSearch.FindAllButton.Enabled = False
                GridSearch.FindFirst.Select()
            End If

            GridSearch.Visible = True
            GridSearch.TextA1.Select()

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub


    Private Sub DeleteSelButton_Click(sender As Object, e As EventArgs) Handles DeleteSelButton.Click
        Dim IntID As Integer    ' integer form of symptom ID
        Dim msg As String       ' error message string

        Try
            ErrorFlag = 1044
            If SelLst.SelectedIndex <> -1 Then
                IntID = GetSymptomID(SelSympData(SelLst.SelectedIndex).SelSympData.ToString)   ' Get symptom ID of selected item.
                SelSympData.RemoveAt(SelLst.SelectedIndex) ' Remove selected item from SelLst.
                SortAndUpdateSelLst()
                If SelLst.Items.Count = 0 Then
                    DeleteSelButton.Enabled = False
                End If

                For SelPtr = 0 To Lst1.SelectedItems.Count - 1
                    If GetSymptomID(CombSympData(Lst1.SelectedIndices(SelPtr)).SympData.ToString) = IntID Then

                        Lst1.SetSelected(Lst1.SelectedIndices(SelPtr), False) 'Deselect the string just deleted.
                        Lst1_Click(sender, e)
                        Exit For
                    End If
                Next SelPtr

                If (SelLst.Items.Count = 0) And (MustBox.Text = "") Then
                    DisableFilePrintPreview()
                    DisableFilePrint()
                    PrescribeButton.Enabled = False
                End If

                If Lst1.SelectedIndex = -1 Then ' Item was de-selected
                    If Lst1.SelectedIndex - 1 Then  ' If no items are selected
                        SelectRemedy.Enabled = False    ' Disable Select Remedy Button
                        MustUse.Enabled = False         ' Disable Must Use button
                        ShowSimilarSymptoms.Enabled = False ' Disable Show Similar Symptoms Button
                    End If
                End If
                Dirty = True
                End If

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
        Buttons:=vbOKOnly + vbCritical,
        Title:="ERROR!")
        End Try
    End Sub
    Public Sub SortAndUpdateSelLst()
        ' Reads text and data for items currently in SelLst into local arrays; sorts them by symptom text (disregards prepended weight), and writes them back to SelLst.

        Dim descStrings As String()     ' Data read from SelLst descriptions minus prepended weight and tab character
        Dim locSelSympData As String()  ' local veraion of SelSympData.SelSympDesc
        Dim locSelSympDesc As String()  ' local version of SelSympData.SelSympData
        Dim locWeights As String()      ' Weight from SelLst descriptions
        Dim minIdx As Integer           ' Index of minimum compare string, used in sort
        Dim msg As String       ' error message string
        Dim ptr, ptr2 As Integer  ' loop pointers
        Dim selLstSize As Integer   ' Number of items in SelLst
        Dim temp1, temp2, temp3     ' temp strings

        Try
            ErrorFlag = 1051
            selLstSize = SelSympData.Count

            ReDim descStrings(selLstSize)
            ReDim locSelSympData(selLstSize)
            ReDim locSelSympDesc(selLstSize)
            ReDim locWeights(selLstSize)

            ' Copy SelLst item(s) to local array.
            For ptr = 0 To selLstSize - 1
                locSelSympDesc(ptr) = SelSympData(ptr).SelSympDesc.ToString
                locSelSympData(ptr) = SelSympData(ptr).SelSympData.ToString
                locWeights(ptr) = Strings.Left(locSelSympDesc(ptr), 1) ' Copy weight character (numeral) to local array
                descStrings(ptr) = Mid(locSelSympDesc(ptr), 3)    ' Remove weight and tab character
            Next ptr

            ' Need to sort SelSympData() items so list item number will match SelLst element number.
            ' NOTE:  Using the ListBox "Sort" option won't work here because it only sorts the display, not the object.
            ' locSelSympData(), locWeights(), and locSelSympData() will be sorted along with descStrings
            ' so they can be added to the correct symptom descriptions.
            For ptr = 0 To selLstSize - 2
                minIdx = ptr
                For ptr2 = ptr + 1 To selLstSize - 1
                    If descStrings(ptr2) < descStrings(minIdx) Then
                        minIdx = ptr2
                    End If
                Next
                If minIdx <> ptr Then
                    temp1 = descStrings(ptr)
                    temp2 = locWeights(ptr)
                    temp3 = locSelSympData(ptr)
                    descStrings(ptr) = descStrings(minIdx)
                    locWeights(ptr) = locWeights(minIdx)
                    locSelSympData(ptr) = locSelSympData(minIdx)
                    descStrings(minIdx) = temp1
                    locWeights(minIdx) = temp2
                    locSelSympData(minIdx) = temp3
                End If
            Next

            ' Put the Desc strings back together and write them to locSelSympDesc().
            SelSympData.Clear()

            For ptr = 0 To selLstSize - 1
                locSelSympDesc(ptr) = locWeights(ptr) + vbTab + descStrings(ptr)
                SelSympData.Add(New SelSympDat(locSelSympDesc(ptr), locSelSympData(ptr)))
            Next

            ' Now write the sorted items back to SelLst.
            SelLst_UpdateDataSource()
            SelLst.Refresh()

        Catch
            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
        Buttons:=vbOKOnly + vbCritical,
        Title:="ERROR!")
        End Try
    End Sub
    Private Sub PrescribeButton_Click(sender As Object, e As EventArgs) Handles PrescribeButton.Click
        Dim AllRem As Integer   ' remedies in symptom list
        Dim DupRem As Boolean   ' more than one symptom for rem flag
        Dim FAWESym As Single   ' normalized auto-weighted equal weight symptom count from SymCntFile
        Dim FAWSym As Single   ' normalized auto-weighted "default" symptom count from SymCntFile
        Dim Found As Boolean    ' item was found in file
        Dim FRemID As Integer   ' remedy ID from SymCntFile
        Dim FSym As Single     ' symptom count from SymCntFile
        Dim FWSym As Single    ' weighted symptom count from SymCntFile
        Dim HitCount(1000) As Integer    ' # hits for each remedy
        Dim Idx As Integer      ' string index
        Dim Largest As Single     ' sort value
        Dim LID As Integer      ' largest ID for sort
        Dim LineDat As String   ' line read from file
        Dim locRemString As String  ' string to display in PrescRem ListBox1
        Dim msg As String       ' error message string
        Dim MustCount As Integer    ' number of "must use" remedies
        Dim MustRem(30) As Integer  ' array of "must use" remedies
        Dim NormMult As Single           ' normalization multiplier
        Dim PickCount As Integer    ' # of ID's in PickRem array
        Dim PickID(1000) As Integer  ' remedy ID's for PickRem's
        Dim PickPtr As Integer  ' pointer to PickID array
        Dim PickRem(1000) As Single    ' array used to pick remedies
        Dim PickStart As Integer    ' starting point in PickID array
        Dim Ptr1, Ptr2 As Integer   ' sort pointers
        Dim RecLen As Integer       ' index file record length
        Dim RemCount As Integer     ' total number of remedy ID's in AllRem list
        Dim RemPtr As Integer       ' remedy pointer
        Dim S As String         ' selection & remedy string
        Dim SelData As String   ' string form of SelSympData
        Dim SelPtr As Integer   ' pointer to selected symptom array
        Dim StrPtr As Integer   ' string pointer
        Dim SympWeight As Integer   ' weight of symptom
        Dim T1 As Single          ' sort swapper
        Dim T2, T3 As Integer   ' sort swappers
        Dim Temp As String      ' substring
        Dim UserWeight As Integer   ' user-supplied weight factor, 1-6

        Try
            ErrorFlag = 1045
            Cursor = Cursors.WaitCursor    ' Change cursor to hourglass.
            MustCount = 0
            PickCount = 0
            RemPtr = 0
            For Ptr1 = 0 To 999
                HitCount(Ptr1) = 0
            Next Ptr1

            MinNumb = Int(MinNum.Text)
            '   Load "must use" array from "must" list.
            If (MustBox.Text <> "") Then    ' If there is a "must use" symptom.
                ErrorFlag = 1151
                S = MustBox.Text
                UserWeight = Val(S)
                StrPtr = 1
                While Mid(MustData, StrPtr, 1) <> "," And StrPtr < Len(MustData)
                    StrPtr = StrPtr + 1 ' Skip symptom ID.
                End While
                StrPtr = StrPtr + 1     ' Now pointing to Remedy ID
                While StrPtr < Len(MustData)
                    MustRem(MustCount) = Val(Mid(MustData, StrPtr))
                    While Mid(MustData, StrPtr, 1) <> "," And StrPtr < Len(MustData)
                        StrPtr = StrPtr + 1 ' Find next number.
                    End While
                    StrPtr = StrPtr + 1 ' Now pointing to symptom number within the remedy; need to look for any "*" after this number

                    SympWeight = 1

                    If (SelectionMode = 1) Then     ' Get weighting factors.
                        If Strings.Right(MustData, 2) = "**" Then
                            SympWeight = 4
                        ElseIf Strings.Right(MustData, 1) = "*" Then
                            SympWeight = 2
                        End If
                    End If
                    MustCount = MustCount + 1

                    PickID(RemPtr) = MustRem(RemPtr)
                    PickRem(RemPtr) = SympWeight * UserWeight
                    HitCount(RemPtr) = HitCount(RemPtr) + 1
                    RemPtr = RemPtr + 1
                    While (Mid(MustData, StrPtr, 1) <> "," And StrPtr < Len(MustData))
                        StrPtr = StrPtr + 1 ' Find next number.
                    End While
                    StrPtr = StrPtr + 1
                End While
                PickCount = MustCount
            End If

            RemCount = 0
            RemPtr = 0
            If (SelLst.Items.Count <> 0) Then
                For SelPtr = 0 To SelLst.Items.Count - 1
                    ErrorFlag = 1046
                    S = SelSympData(SelPtr).SelSympDesc.ToString
                    UserWeight = Val(S)
                    SelData = SelSympData(SelPtr).SelSympData.ToString
                    StrPtr = 1
                    While (Mid(SelData, StrPtr, 1) <> "," And StrPtr < Len(SelData))
                        StrPtr = StrPtr + 1 ' Find next number.
                    End While
                    StrPtr = StrPtr + 1
                    While (StrPtr < Len(SelData))
                        RemCount = RemCount + 1
                        AllRem = Val(Mid(SelData, StrPtr))
                        While (Mid(SelData, StrPtr, 1) <> "," And StrPtr < Len(SelData))
                            StrPtr = StrPtr + 1 ' Find next number.
                        End While
                        StrPtr = StrPtr + 1
                        Temp = Mid(SelData, StrPtr)
                        Idx = InStr(Temp, ",")
                        If (Idx = Len(Temp)) Then Idx = Idx - 1
                        Temp = Mid(Temp, 1, Idx)    ' Get rid of the comma or new-line.
                        While (Mid(SelData, StrPtr, 1) <> "," And StrPtr < Len(SelData))
                            StrPtr = StrPtr + 1 ' Find next number.
                        End While
                        StrPtr = StrPtr + 1
                        SympWeight = 1
                        If (SelectionMode = 1) Then 'Get weighting factors.
                            If Idx > 2 Then
                                If Mid(Temp, Idx - 2, 2) = "**" Then _
                                    SympWeight = 4
                            End If
                            If Idx >= 1 Then
                                If Idx > 1 And Mid(Temp, Idx - 1, 1) = "*" Then SympWeight = 2
                            End If
                        End If

                        If (MustCount = 0) Then 'If no "Must use" remedy
                            DupRem = False
                            For PickPtr = 0 To PickCount - 1
                                If (PickID(PickPtr) = AllRem) Then  'If PickID matches the current remedy ID from string.
                                    DupRem = True
                                    PickRem(PickPtr) = PickRem(PickPtr) + SympWeight * UserWeight
                                    HitCount(PickPtr) = HitCount(PickPtr) + 1
                                End If
                            Next PickPtr

                            If (Not DupRem) Then
                                PickID(PickCount) = AllRem
                                PickRem(PickCount) = SympWeight * UserWeight
                                HitCount(PickCount) = HitCount(PickCount) + 1
                                PickCount = PickCount + 1
                            End If
                        Else    'Use only "Must use" remedies.
                            For PickPtr = 0 To MustCount - 1
                                If (PickID(PickPtr) = AllRem) Then
                                    PickRem(PickPtr) = PickRem(PickPtr) + (SympWeight * UserWeight)
                                    HitCount(PickPtr) = HitCount(PickPtr) + 1
                                End If
                            Next PickPtr
                        End If
                    End While
                Next SelPtr
            End If

            If (SelectionMode <> 3) Then    ' Need to get total symptoms per remedy from file.
                ErrorFlag = 1047
                If (MustCount = 0) Then
                    PickStart = 1
                Else
                    PickStart = 0
                End If
                For PickPtr = 0 To PickCount - 1
                    Found = False
                    ' Remedies are numbered consecutively starting at 1.  RemData array is 0-based, so need to subtract 1.
                    Temp = RemData(PickID(PickPtr) - 1) ' Data from RemPresc.dat:  4-character digits separated by commas
                    FRemID = Val(Strings.Left(Temp, 4))
                    Temp = Temp.Substring(5)
                    FSym = Val(Strings.Left(Temp, 4))
                    Temp = Temp.Substring(5)
                    FWSym = Val(Strings.Left(Temp, 4))
                    Temp = Temp.Substring(5)
                    FAWSym = Val(Strings.Left(Temp, 4))
                    Temp = Temp.Substring(5)
                    FAWESym = Val(Temp)
                    If (FRemID = PickID(PickPtr)) Then
                        Found = True
                        If (SelectionMode = 1) Then
                            If Normalize Then
                                PickRem(PickPtr) = PickRem(PickPtr) * 100 / FAWSym
                            Else
                                PickRem(PickPtr) = PickRem(PickPtr) * 100 / FWSym
                            End If
                        Else
                            If (SelectionMode = 2) Then
                                If Normalize Then
                                    PickRem(PickPtr) = PickRem(PickPtr) * 100 / FAWESym
                                Else
                                    PickRem(PickPtr) = PickRem(PickPtr) * 100 / FSym
                                End If
                            End If
                        End If
                    End If
                Next PickPtr
            End If

            ErrorFlag = 1048
            For Ptr1 = 0 To PickCount - 1
                Largest = -1
                For Ptr2 = Ptr1 To PickCount - 1
                    If (PickRem(Ptr2) >= Largest) Then
                        Largest = PickRem(Ptr2)
                        LID = Ptr2
                    End If
                Next Ptr2
                T1 = PickRem(Ptr1)
                T2 = PickID(Ptr1)
                T3 = HitCount(Ptr1)
                PickRem(Ptr1) = PickRem(LID)
                PickID(Ptr1) = PickID(LID)
                HitCount(Ptr1) = HitCount(LID)
                PickRem(LID) = T1
                PickID(LID) = T2
                HitCount(LID) = T3
            Next Ptr1

            ErrorFlag = 1049
            If Normalize And SelectionMode = 3 Then
                If PickRem(0) <> 0# Then
                    NormMult = 100.0# / PickRem(0)
                Else
                    NormMult = 0#
                End If
            End If

            ErrorFlag = 1050
            RecLen = 8

            DisplayCount = 0
            If (PickCount > 1000) Then PickCount = 1000   'Don't overflow array.
            DisplayCount = 0
            SelRemData.Clear()    ' Remove items from any previous Prescribe
            For RemPtr = 0 To PickCount - 1
                If (HitCount(RemPtr) >= MinNumb) Then
                    If (PickID(RemPtr) > file_size Or PickID(RemPtr) < 0) Then
                        ErrorFlag = 1053
                        ' Force a trap
                        Throw New System.Exception("An exception has occurred.")
                    End If

                    ErrorFlag = 1055
                    LineDat = RemText(PickID(RemPtr) - 1) ' Remedy text from RemPresc.dat

                    ErrorFlag = 1056
                    If Normalize And SelectionMode = 3 Then
                        Pres(DisplayCount) = Format((PickRem(RemPtr) * NormMult), "##0.00") + Chr(9) + Chr(9) + LineDat
                    Else
                        Pres(DisplayCount) = Format(PickRem(RemPtr), "##0.00") + Chr(9) + Chr(9) + LineDat
                    End If
                    locRemString = Pres(DisplayCount)
                    If InStr(locRemString, "\") > 0 Then locRemString = Strings.Left(locRemString, InStr(locRemString, "\") - 1)
                    SelRemData.Add(New SelRemDat(locRemString, PickID(RemPtr).ToString))
                    DisplayCount = DisplayCount + 1
                End If
            Next RemPtr

            Cursor = Cursors.Default    ' Change cursor to default.

            If DisplayCount > 0 Then
                If PrescRem.Visible Then PrescRem.Close()   ' Need to re-initialize in case it is open.

                ' Set up data source for SelLst
                PrescRem.ListBox1.DataSource = Form1.SelRemData
                PrescRem.ListBox1.ValueMember = "SelRemData"
                PrescRem.ListBox1.DisplayMember = "SelRemDesc"
                PrescRem.ListBox1.ClearSelected()

                PrescRem.Top = 1500
                PrescRem.Left = 250
                PrescRem.Show()
            Else
                MsgBox(Prompt:="No remedy contains the Minimum Number of symptoms.",,
                    vbOKOnly + vbInformation)
            End If

        Catch
            Cursor = Cursors.Default    ' Change cursor to default.
            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
            Buttons:=vbOKOnly + vbCritical,
            Title:="ERROR!")
        End Try
    End Sub


    Private Sub ShowSimilarSymptoms_Click(sender As Object, e As EventArgs) Handles ShowSimilarSymptoms.Click

        Handle_ShowSimilarSymptoms()
    End Sub

    Private Sub Handle_ShowSimilarSymptoms()
        Dim ErrorFlag As Integer    ' make it local to preserve it after calls
        Dim msg As String       ' error pop-up string
        Dim ptr1, ptr2 As Integer      ' loop pointers

        Try
            Form1.SympSize = 1
            Form1.SMaxSympPtr = Form1.SCombSympPtr

            ErrorFlag = 1057
            QuestProgress.STDCount = 0  'Initialize symptom-to-display count
            CalledFromShowSymSymps = True

            ' Copy seartxtidx.dat to 2 string arrays to make word searching easier.
            ReDim QuestProgress.SearWords(Form1.Sfile_size / 2)
            ReDim QuestProgress.SearPointers(Form1.Sfile_size / 2)
            ptr2 = 0
            For ptr1 = 0 To Form1.Sfile_size - 1 Step 2
                QuestProgress.SearWords(ptr2) = Form1.GlobalSearchData(ptr1)
                QuestProgress.SearPointers(ptr2) = Form1.GlobalSearchData(ptr1 + 1)
                ptr2 += 1
            Next

            FindAllCaller = 4   ' Tell FindAll it's being called by Show Similar Symptoms
            ReDim QuestProgress.WordArray(1)

            ' Pre-select options to require a 60% match of words or synonyms.
            ' QuestProgress.FindWords also requires symptoms to contain match of 1st word or its synonyms.
            QuestProgress.PercentageRequired = 0.6

            '   Assign weights to the symptoms based on check boxes on
            '   page 1.
            Call QuestProgress.SymParse(TextBox1.Text, Form1.QuestWeight)

            If QuestProgress.STDCount > 0 Then
                Call QuestProgress.AddList()  ' Show FindAll, populate list with found symptoms.
            Else
                ' Handle case where no matches were found.
                MsgBox("No symptom matches were found.", vbOKOnly + vbInformation)
            End If
        Catch
            Cursor = Cursors.Default    ' Change cursor to default.
            msg = "Oops, something went wrong; please try again. (" + Str(ErrorFlag) + ")"
            Dim unused = MsgBox(Prompt:=msg,
                Buttons:=vbOKOnly + vbCritical,
                Title:="ERROR!")
        End Try
    End Sub

    Private Sub SelLst_Click(sender As Object, e As EventArgs) Handles SelLst.Click
        Dim str As String   ' String for Textbox1 Text

        If MustBoxSelected Then
            MustBox.Select(0, 0)
            MustBoxSelected = False
        End If
        str = SelSympData(SelLst.SelectedIndex).SelSympDesc.ToString()
        TextBox1.Text = Strings.Mid(str, InStr(str, vbTab) + 1)    ' Display symptom in text box; can be used for "find similar symptoms"

    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles buttonDefault.CheckedChanged
        OptionsDefault()
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles buttonEqualWeight.CheckedChanged
        OptionsEqualWeight()
    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles buttonStraightCount.CheckedChanged
        OptionsNormalize()
    End Sub


    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles Holostic.CheckedChanged
        If Holostic.Checked Then
            OptionsAutoWeight()
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged_1(sender As Object, e As EventArgs) Handles boxNormalize.CheckedChanged
        OptionsNormalize()
    End Sub

    Private Sub OptionsMenuItem_Click(sender As Object, e As EventArgs) Handles OptionsMenuItem.Click
        OptionsSetPreferences()
    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        OptionsSetPreferences()
    End Sub

    Private Sub MustBox_Click(sender As Object, e As EventArgs) Handles MustBox.Click
        Dim str As String   ' Message string for TextBox1

        If Len(MustBox.Text) > 0 Then
            If MustBoxSelected Then
                MustBox.Select(0, 0)
                MustBoxSelected = False
            Else
                ' Reset any selected item in SelLst (so MustBox behaves like an extension to SelLst)
                If SelLst.SelectedIndex <> -1 Then
                    SelLst.SetSelected(SelLst.SelectedIndex, False)
                End If
                MustBox.SelectAll()
                MustBoxSelected = True
                str = MustBox.Text
                TextBox1.Text = Strings.Mid(str, InStr(str, vbTab) + 1)    ' Display symptom in text box; can be used for "find similar symptoms"
            End If
        End If
    End Sub

    Private Sub Weight_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Weight.SelectedIndexChanged
        Dim itemCount As Integer    ' number of items in SelLst
        Dim modPtr As Integer   ' points to SelLst item to modify
        Dim SelPtr As Integer   ' points to item in SelLst
        Dim locSelSympDesc As String() ' SelLst items text
        Dim locSelSympData As String() ' SelLst items data
        Dim tempSympText As String  ' text of symptom item that needs weight value updated

        ' Synchronize Weight in OneSymp and FindAll
        OneSymp.Weight.Text = Weight.Text
        FindAll.Weight.Text = Weight.Text

        ' If an item in MustBox or SelLst is selected, need to update its weight value.
        If MustBoxSelected Then
            ' Change weight in the symptom text to the new value.
            tempSympText = MustBox.Text
            ' Remove the weight from the text
            tempSympText = Mid(tempSympText, 2, Len(tempSympText))
            MustBox.Text = Weight.Text + tempSympText
            ' De-select MustBox
            MustBox.Select(0, 0)
            MustBoxSelected = False
        ElseIf SelLst.SelectedIndex <> -1 Then  ' Item is in SelLst
            ' There doesn't appear to be any mechanism for updating a single member of a ListBox; so need to copy contents of list
            ' and its data to 2 arrays, update the weight in the selected item, delete all items in SelLst, then copy back the
            ' arrays with the updated weight item.
            itemCount = SelLst.Items.Count
            ReDim locSelSympDesc(itemCount)
            ReDim locSelSympData(itemCount)
            modPtr = SelLst.SelectedIndex   ' save off index of item to modify

            ' Copy SelLst item text and data to local arrays
            For SelPtr = 0 To itemCount - 1
                locSelSympDesc(SelPtr) = SelLst.Items(SelPtr).SelSympDesc.ToString
                locSelSympData(SelPtr) = SelLst.Items(SelPtr).SelSympData.ToString
            Next

            ' Update the weight in the selected item
            tempSympText = locSelSympDesc(modPtr)
            tempSympText = Mid(tempSympText, 2, Len(tempSympText))
            locSelSympDesc(modPtr) = Weight.Text + tempSympText

            ' Clear out the list; items will be added back with (updated) locSelLst...
            SelSympData.Clear()
            SelLst_UpdateDataSource()

            ' Copy arrays containing the updated text item back to SelLst
            For SelPtr = 0 To itemCount - 1
                SelSympData.Add(New SelSympDat(locSelSympDesc(SelPtr), locSelSympData(SelPtr)))
            Next
            SelLst_UpdateDataSource()
            SelLst.Refresh()
            SelLst.SetSelected(modPtr, True)    ' Re-select the updated item.
        End If
        Dirty = True    ' In case data was read from a case file, need to update it with the new symptom weight.

    End Sub


    Private Sub ToolTip1_Popup(sender As Object, e As PopupEventArgs) Handles ToolTip1.Popup
        TooltipText.Text = ToolTip1.GetToolTip(e.AssociatedControl)
    End Sub

    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click
        Dim myProcess As New Process
        myProcess.StartInfo.FileName = AppContext.BaseDirectory + MAN_DIR + "Dr Hahnemanns Assistant User Guide Version 3.pdf"
        myProcess.StartInfo.UseShellExecute = True
        myProcess.StartInfo.RedirectStandardOutput = False
        myProcess.Start()
        myProcess.Dispose()
    End Sub

    Private Sub AcceptDisc_Click(sender As Object, e As EventArgs) Handles AcceptDisc.Click
        ' Hides the Disclaimer
        HideDisclaimer()

        If ShowIntroduction Then
            VisibleIntroduction()
        End If
    End Sub

    Private Sub IntroContinue_Click(sender As Object, e As EventArgs) Handles IntroContinue.Click
        ' Hides the introduction
        HideIntroduction()
    End Sub
    Private Sub ExitDisc_Click(sender As Object, e As EventArgs) Handles ExitDisc.Click
        FileExit()
    End Sub
    Private Sub IntroExit_Click(sender As Object, e As EventArgs) Handles IntroExit.Click
        FileExit()
    End Sub
    Private Sub DontShowDisc_CheckedChanged(sender As Object, e As EventArgs) Handles DontShowDisc.CheckedChanged
        If DontShowDisc.Checked Then
            SaveSetting("Hahnasst", "Startup",
                   "ShowDisclaimer", False)
        End If
    End Sub
    Private Sub DontShowIntro_CheckedChanged(sender As Object, e As EventArgs) Handles DontShowIntro.CheckedChanged
        If DontShowIntro.Checked Then
            SaveSetting("Hahnasst", "Startup",
                   "ShowIntroduction", False)
        End If
    End Sub

    Private Sub HideDisclaimer()
        DiscText.Visible = False
        AcceptDisc.Visible = False
        ExitDisc.Visible = False
        DontShowDisc.Visible = False
    End Sub

    Private Sub HideIntroduction()
        IntroText.Visible = False
        IntroContinue.Visible = False
        IntroExit.Visible = False
        DontShowIntro.Visible = False
    End Sub

    Private Sub VisibleDisclaimer()
        DiscText.Visible = True
        AcceptDisc.Visible = True
        ExitDisc.Visible = True
        DontShowDisc.Visible = True
    End Sub

    Private Sub VisibleIntroduction()
        IntroText.Visible = True
        IntroContinue.Visible = True
        IntroExit.Visible = True
        DontShowIntro.Visible = True
    End Sub

    Private Sub StatusStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles StatusStrip1.ItemClicked

    End Sub

End Class

Public Class CombSympDat
    Private mySympData As String
    Private mySympDesc As String

    Public Sub New(ByVal strSympDesc As String, ByVal strSympData As String)
        Me.mySympData = strSympData
        Me.mySympDesc = strSympDesc
    End Sub

    Public ReadOnly Property SympData() As String
        Get
            Return mySympData
        End Get
    End Property

    Public ReadOnly Property SympDesc() As String
        Get
            Return mySympDesc
        End Get
    End Property

End Class

Public Class SelRemDat
    Private mySelRemDesc As String
    Private mySelRemData As String

    Public Sub New(ByVal strSelRemDesc As String, ByVal strSelRemData As String)
        Me.mySelRemDesc = strSelRemDesc
        Me.mySelRemData = strSelRemData
    End Sub

    Public ReadOnly Property SelRemData() As String
        Get
            Return mySelRemData
        End Get
    End Property

    Public ReadOnly Property SelRemDesc() As String
        Get
            Return mySelRemDesc
        End Get
    End Property

End Class
Public Class SelSympDat
    Private mySelSympDesc As String
    Private mySelSympData As String

    Public Sub New(ByVal strSelSympDesc As String, ByVal strSelSympData As String)
        Me.mySelSympDesc = strSelSympDesc
        Me.mySelSympData = strSelSympData
    End Sub

    Public ReadOnly Property SelSympData() As String
        Get
            Return mySelSympData
        End Get
    End Property

    Public ReadOnly Property SelSympDesc() As String
        Get
            Return mySelSympDesc
        End Get
    End Property

End Class