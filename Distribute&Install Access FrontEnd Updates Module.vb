
'*********************************************************************************
'Code by: Edgar D Hernandez, LinkedIn: https://www.linkedin.com/in/edgardhernandez1/
'
'This code can be run at start up and is meant to be used to facilitates the proccess
'of pushing and receiving front-end database updates, as long as the front-end is or
'can be placed on a shared-drive location.
'
'The code checks to see what version the user has.
'Displays the version number and update description of the version(s) missed.
'Enables the button(s) the user is allowed to use based on if they need to update or not
'Changes the background color of the form based on if and type of update needed.
'
'Things to keep in mind:
'To display the update message to the users run the UpdateUI() function once
'To display the push/distribute updates settings form, run the UpdateUI() twice
'Update the function fSharedrivePath() string to the file path of your front end database location
'Also the following functions can be customized: myForm(), myTableFE(), and myTableBE()
'Make sure that the table myTableBE is a linked table and that the myTableFE remains as a local table in the user's front end copy
'When pushing updates to the user make sure you are doing so from your desktop
'
'*********************************************************************************

Private Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare PtrSafe Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Function fSharedrivePath() As String
    'FRONTEND DATABASE SHAREPOINT LOCATION
    fSharedrivePath = "C:\Users\" & fSystemUserName & "\Documents"
End Function

Public Function myForm() As Variant
    'FORM NAME
    myForm = "frm_Version_Update"
End Function

Public Function myTableFE() As Variant
    'FRONT END VERSION TABLE
    myTableFE = "USysVersionControl"
End Function

Public Function myTableBE() As Variant
    'BACK END VERSION TABLE
    myTableBE = "USysVersionServer"
End Function

Public Function UpdateUI()

'<!-- /*------------------------------*/ //-->
' PROCEDURE VARIABLES
'<!-- /*------------------------------*/ //-->

'TWIP (TWentieth of an InchPoint)
'THERE ARE 72 TWIPS PER POINT (1440/20)
Dim twipsPerInch, qtrInch As Integer
twipsPerInch = 1440
qtrInch = 1440 / 4

'FONT WEIGHT
Dim fontWeight_Normal, fontWeight_Bold As Integer
fontWeight_Normal = 400
fontWeight_Bold = 700

'VERSION NUMBER FORMAT
Dim newVersionNumber As String
newVersionNumber = Right(Year(Now()), 2) & "." & Month(Now()) & "." & Day(Now()) & "." & Hour(Now()) & Minute(Now())

'FIELDS
On Error Resume Next
    Dim tbl As TableDef
    Dim fld, fld1, fld2, fld3, fld4, fld5, fld6 As Field

'<!-- /*------------------------------*/ //-->
' CHECK TO SEE IF FRONTEND TABLE ALREADY EXISTS
'<!-- /*------------------------------*/ //-->

DoCmd.SetWarnings False
Dim strTableServer As String
Dim rstTableServer As DAO.Recordset

Set rstTableServer = CurrentDb.OpenRecordset("SELECT Count(IIf([MSysNavPaneObjectIDs]![Name]='" & myTableFE & "',1)) AS Expr1 " & vbCrLf & _
    "FROM MSysNavPaneObjectIDs;")

    strTableServer = rstTableServer!Expr1
    rstTableServer.Close
Set rstTableServer = Nothing
DoCmd.SetWarnings True

'<!-- /*------------------------------*/ //-->
' CHECK TO SEE IF BACKEND TABLE ALREADY EXISTS
'<!-- /*------------------------------*/ //-->

DoCmd.SetWarnings False
Dim strTableControl As String
Dim rstTableControl As DAO.Recordset

Set rstTableControl = CurrentDb.OpenRecordset("SELECT Count(IIf([MSysNavPaneObjectIDs]![Name]='" & myTableBE & "',1)) AS Expr1 " & vbCrLf & _
    "FROM MSysNavPaneObjectIDs;")

    strTableControl = rstTableControl!Expr1
    rstTableControl.Close
Set rstTableControl = Nothing
DoCmd.SetWarnings True

'<!-- /*------------------------------*/ //-->
' CREATE FRONTEND TABLE IF IT DOESNT EXISTS
'<!-- /*------------------------------*/ //-->

If strTableServer = 0 Then

    DoCmd.SetWarnings False
    DoCmd.Close acTable, myTableFE, acSaveYes
    DoCmd.DeleteObject acTable = acDefault, myTableFE

    Set tbl = CurrentDb.CreateTableDef(myTableFE)

    'DESIGN THE FIELDS
    Set fld = tbl.CreateField("ID", dbLong)
        fld.Attributes = dbAutoIncrField
    Set fld1 = tbl.CreateField("VersionNumber", dbText)

    'CREATE THE FIELDS
    tbl.Fields.Append fld
    tbl.Fields.Append fld1

    'CREATE AND DISPLAY THE TABLE
    CurrentDb.TableDefs.Append tbl
    CurrentDb.TableDefs.Refresh

    Application.SetHiddenAttribute acTable, myTableFE, True

    DoCmd.RunSQL "INSERT INTO " & myTableFE & " ( VersionNumber ) SELECT '0.0.0.0' AS Initial_Version;"
    DoCmd.SetWarnings True

End If

'<!-- /*------------------------------*/ //-->
' CREATE BACKEND TABLE IF IT DOESNT EXISTS
'<!-- /*------------------------------*/ //-->

If strTableControl = 0 Then

    DoCmd.SetWarnings False
    DoCmd.Close acTable, myTableBE, acSaveYes
    DoCmd.DeleteObject acTable = acDefault, myTableBE

    Set tbl = CurrentDb.CreateTableDef(myTableBE)

    'DESIGN THE FIELDS
    Set fld = tbl.CreateField("ID", dbLong)
        fld.Attributes = dbAutoIncrField
    Set fld1 = tbl.CreateField("Application_Name", dbText)
    Set fld2 = tbl.CreateField("Version_Number", dbText)
    Set fld3 = tbl.CreateField("Updates", dbMemo)
    Set fld4 = tbl.CreateField("DTG", dbDate)
    Set fld5 = tbl.CreateField("Updated_By", dbText)
    Set fld6 = tbl.CreateField("Mandatory", dbBoolean)

    'CREATE THE FIELDS
    tbl.Fields.Append fld
    tbl.Fields.Append fld1
    tbl.Fields.Append fld2
    tbl.Fields.Append fld3
    tbl.Fields.Append fld4
    tbl.Fields.Append fld5
    tbl.Fields.Append fld6

    'CREATE AND DISPLAY THE TABLE
    CurrentDb.TableDefs.Append tbl
    CurrentDb.TableDefs.Refresh

    Application.SetHiddenAttribute acTable, myTableBE, True

    DoCmd.RunSQL "INSERT INTO " & myTableBE & " ( Application_Name, Version_Number, Updates, DTG, Updated_By, Mandatory ) SELECT fApplicationName() AS Application_Name, '0.0.0.0' AS Initial_Version, 'Initial Version' AS Updates, Now() AS DTG, fSystemUserName() AS Updated_By, False AS Mandatory;"
    DoCmd.SetWarnings True

End If

'<!-- /*------------------------------*/ //-->
' CHECK TO SEE IF THE FORM ALREADY EXISTS
'<!-- /*------------------------------*/ //-->

DoCmd.SetWarnings False
Dim strFormCheck As String
Dim rstForm As DAO.Recordset

Set rstForm = CurrentDb.OpenRecordset("SELECT Count(IIf([MSysNavPaneObjectIDs]![Name]='" & myForm & "',1)) AS Expr1 " & vbCrLf & _
    "FROM MSysNavPaneObjectIDs;")

    strFormCheck = rstForm!Expr1
    rstForm.Close
Set rstForm = Nothing
DoCmd.SetWarnings True

On Error Resume Next

'<!-- /*------------------------------*/ //-->
' CREATE FORM IF IT DOESNT EXISTS
'<!-- /*------------------------------*/ //-->

If strFormCheck = 0 Then
'If strFormCheck <= 1 Then

    DoCmd.SetWarnings False

    Dim defaultFormName As Form
    Set defaultFormName = CreateForm

    DoCmd.Close acForm, defaultFormName, acSaveNo
    DoCmd.Close acForm, myForm, acSaveNo
    DoCmd.DeleteObject acForm = acDefault, myForm

    Dim ctlNameArray(), ctlTypeArray() As Variant
        ctlNameArray = Array("Version_Number", "Version_Current", "Updates", "DTG", "Updated_By", "msg_info", "Mandatory", "Update", "UpdateClose", "Close")
        ctlTypeArray = Array(acTextBox, acTextBox, acTextBox, acTextBox, acTextBox, acTextBox, acCheckBox, acCommandButton, acCommandButton, acCommandButton)

    Dim c As Integer
    Dim ctlText, ctlLabel As Control
    For c = 0 To UBound(ctlNameArray)
        Set ctlText = CreateControl(defaultFormName.Name, ctlTypeArray(c), , "", "", qtrInch * 10, qtrInch * c)
        If c < 5 Then
            Set ctlLabel = CreateControl(defaultFormName.Name, acLabel, , "", "Label_" & ctlNameArray(c), 100, qtrInch * c)
            ctlLabel.Name = "Label_" & ctlNameArray(c)
        End If
            ctlText.Name = ctlNameArray(c)
    Next c

    Dim dfm As String
        dfm = defaultFormName.Name

    DoCmd.Restore
    DoCmd.Close acForm, defaultFormName.Name, acSaveYes
    DoCmd.Rename myForm, acForm, dfm

'<!-- /*------------------------------*/ //-->
' FOMART FORM ON DESIGN VIEW
'<!-- /*------------------------------*/ //-->

    DoCmd.OpenForm myForm, acDesign, , , , acWindowNormal
        Forms(myForm).PopUp = True
        Forms(myForm).AllowDatasheetView = False
        Forms(myForm).AllowLayoutView = False
        Forms(myForm).BorderStyle = 3

        Forms(myForm).Width = twipsPerInch * 5
        Forms(myForm).Detail.Height = twipsPerInch * 3.5

        Forms(myForm).Modal = True
        Forms(myForm).RecordSelectors = False
        Forms(myForm).NavigationButtons = False
        Forms(myForm).ScrollBars = 0
        Forms(myForm).ControlBox = False
        Forms(myForm).Moveable = True
    DoCmd.Close acForm, myForm, acSaveYes

    DoCmd.SetWarnings False

End If

'<!-- /*------------------------------*/ //-->
' FOMART FORM ON NORMAL VIEW
'<!-- /*------------------------------*/ //-->

DoCmd.OpenForm myForm, acNormal, , , , acWindowNormal

Dim formNameArray As Variant
    formNameArray = Array(Forms(myForm))
Set FormName = formNameArray(0)
With FormName

    FormName.Detail.BackColor = RGB(37, 64, 97)

    FormName.Caption = "Push Software Update(s)"
    FormName.[Mandatory] = False

'<!-- /*------------------------------*/ //-->
' LABEL FORMAT VARIABLES
'<!-- /*------------------------------*/ //-->

    'LABEL NAME
    Dim labelNameArray() As Variant
        labelNameArray = Array(FormName.Label_Version_Number, FormName.Label_Version_Current, FormName.Label_Updates, FormName.Label_DTG, FormName.Label_Updated_By)

    'LABEL DIMENSIONS
    Dim labelWidth(), labelHeight() As Variant
        labelWidth = Array(1.375, 1.375, 1, 1, 1)
        labelHeight = Array(0.375, 0.375, 0.5, 0.2083, 0.2083)

    'LABEL POSITION
    Dim labelMoveLeft(), labelMoveTop() As Variant
        labelMoveTop = Array(0.125, 0.125, 0.875, 1.75, 2)
        labelMoveLeft = Array(1.375, 3.375, 0.25, 0.25, 0.25)

    'LABEL DISPLAY VALUE
    Dim labelCaption() As Variant
        labelCaption = Array("Version Being Pushed", "Make Updates Mandatory", "Updates Made", "Current Date", "Updated By")

    'LABEL VISIBILITY
    Dim labelVisible As Variant
        labelVisible = Array(True, True, True, True, True)

'<!-- /*------------------------------*/ //-->
' TEXTBOX FORMAT VARIABLES
'<!-- /*------------------------------*/ //-->

    'TEXTBOX NAME
    Dim textboxNameArray() As Variant
        textboxNameArray = Array(FormName.[Version_Number], FormName.[Version_Current], FormName.[Updates], FormName.[DTG], FormName.[Updated_By], FormName.[msg_info])

    'TEXTBOX DIMENSIONS
    Dim textboxWidth(), textboxHeight() As Variant
        textboxWidth = Array(1.375, 1.375, 3.375, 3.375, 3.375, 4.5)
        textboxHeight = Array(0.2083, 0.2083, 0.8354, 0.2083, 0.2083, 0.5)

    'TEXTBOX POSITION
    Dim textboxMoveLeft(), textboxMoveTop() As Variant
        textboxMoveTop = Array(0.5417, 0.5417, 0.875, 1.75, 2, 2.2917)
        textboxMoveLeft = Array(1.375, 3.375, 1.375, 1.375, 1.375, 0.25)

    'TEXTBOX DISPLAY VALUE
    Dim textboxValue()  As Variant
        textboxValue = Array(newVersionNumber, "0.0.0.0", Null, Now(), fSystemUserName(), "Version Update")

    'TEXTBOX VISIBILITY
    Dim textboxVisible, textboxEnabled, textboxLocked As Variant
        textboxVisible = Array(True, False, True, True, True, True)
        textboxEnabled = Array(False, False, True, False, False, False)
        textboxLocked = Array(True, True, False, True, True, True)

'<!-- /*------------------------------*/ //-->
' BUTTON FORMAT VARIABLES
'<!-- /*------------------------------*/ //-->

    'BUTTON NAMES
    Dim buttonNameArray() As Variant
        buttonNameArray = Array(FormName.[Update], FormName.[UpdateClose], FormName.[Close])

    'BUTTON DIMENSIONS
    Dim buttonWidth(), buttonHeight() As Variant
        buttonWidth = Array(1.0833, 1.0833, 1.0833)
        buttonHeight = Array(0.25, 0.25, 0.25)

    'BUTTON POSITION
    Dim buttonMoveLeft(), buttonMoveTop() As Variant
        buttonMoveTop = Array(2.9583, 2.9583, 2.9583)
        buttonMoveLeft = Array(0.25, 1.9583, 3.6667)

    'BUTTON DISPLAY VALUE
    Dim buttonCaption() As Variant
        buttonCaption = Array("Push", "Notify", "Close")

    'BUTTON VISIBILITY
    Dim buttonVisible, buttonEnabled As Variant
        buttonVisible = Array(True, True, True)
        buttonEnabled = Array(True, False, True)
        'buttonEnabled = Array(False, False, True)

'<!-- /*------------------------------*/ //-->
' CHECK FOR UPDATES
'<!-- /*------------------------------*/ //-->

    If FormName.Mandatory.Visible = True Then

        FormName.Caption = "Check for Updates"

        'LABEL DISPLAY VALUE
        labelCaption = Array("Newest Version Available", "Current Version Installed", "What's New?", "Update Date", "Updates By")

        'LABEL VISIBILITY
        labelVisible = Array(True, True, True, True, True)

        'TEXTBOX VISIBILITY
        textboxVisible = Array(True, True, True, True, True)
        textboxEnabled = Array(False, False, True, False, False)
        textboxLocked = Array(True, True, True, True, True)

'<!-- /*------------------------------*/ //-->
' GET THE ID OF THE NEWEST VERSION AVAILABLE
'<!-- /*------------------------------*/ //-->

    On Error Resume Next
        Dim rstMaxIdDTG As DAO.Recordset
        Dim rstData As DAO.Recordset
    
        Set rstMaxIdDTG = CurrentDb.OpenRecordset("SELECT Max([" & myTableBE & "]![DTG]) AS NewestDTG, Max([ID]) AS NewestID FROM " & myTableBE & ";")
        Set rstData = CurrentDb.OpenRecordset("SELECT " & myTableBE & ".Version_Number, " & myTableBE & ".DTG, " & myTableBE & ".[Updated_By], " & myTableBE & ".ID " & vbCrLf & _
            "FROM " & myTableBE & " " & vbCrLf & _
            "WHERE (((" & myTableBE & ".ID)=" & rstMaxIdDTG!NewestID & "));")
    
        If rstMaxIdDTG!NewestDTG = rstData!DTG Then
            FormName.[Version_Number] = rstData!Version_Number
            FormName.[DTG] = rstData!DTG
            FormName.[Updated_By] = rstData![Updated_By]
        End If

        rstData.Close
        Set rstData = Nothing

'<!-- /*------------------------------*/ //-->
' GET THE RECORD ID THAT CORRESPONDS TO THE
' USER'S CURRENT VERSION NUMBER
'<!-- /*------------------------------*/ //-->

        Dim strCVID As String
        Dim rstCurrentVersion As DAO.Recordset
        Set rstCurrentVersion = CurrentDb.OpenRecordset("SELECT " & myTableBE & ".Version_Number, " & myTableBE & ".ID " & vbCrLf & _
            "FROM " & myTableBE & " INNER JOIN " & myTableFE & " ON " & myTableBE & ".Version_Number = " & myTableFE & ".VersionNumber;")
    
        FormName.[Version_Current] = rstCurrentVersion!Version_Number

        strCVID = rstCurrentVersion!ID
        rstCurrentVersion.Close
        Set rstCurrentVersion = Nothing

'<!-- /*------------------------------*/ //-->
' GET THE COUNT OF HOW MANY VERSION UPDATES
' HAVE BEEN MISSED BY THE USER
'<!-- /*------------------------------*/ //-->

        Dim strCount As String
        Dim rstCount As DAO.Recordset

        Set rstCount = CurrentDb.OpenRecordset("SELECT Count(" & myTableBE & ".Version_Number) AS CountOfVersion_Number " & vbCrLf & _
            "FROM " & myTableBE & " " & vbCrLf & _
            "WHERE (((" & myTableBE & ".ID)>" & strCVID & "));")

        strCount = rstCount!CountOfVersion_Number
        rstCount.Close
        Set rstCount = Nothing

'<!-- /*------------------------------*/ //-->
' GET THE VERSION NUMBER AND DISCRIPTION
' OF THE UPDATES MISSED
'<!-- /*------------------------------*/ //-->

        Dim strUpdates As String
        Dim rst As DAO.Recordset
  
        Set rst = CurrentDb.OpenRecordset("SELECT [Version_Number] & ':' & Space(1) & [Updates] AS [New Updates], " & myTableBE & ".[Mandatory], " & myTableBE & ".ID " & vbCrLf & _
            "FROM " & myTableBE & " " & vbCrLf & _
            "WHERE (((" & myTableBE & ".ID)>" & strCVID & "));")

        Do Until rst.EOF
            strUpdates = strUpdates & rst![New Updates] & vbCrLf
            rst.MoveNext
        Loop
        rst.Close
        Set rst = Nothing

        If strCount >= 1 Then
            FormName.[Updates] = strUpdates
        Else
            FormName.[Updates] = "No new updates!"
        End If

'<!-- /*------------------------------*/ //-->
' CHECKS TO SEE IF ANY OF THE MISSED UPDATES WHERE MANDATORY
' AND IF SO, THIS WILL FORCE THEM TO UPDATE
' REGRADLESS IF THE RECENT UPDATE WAS MANDATORY OR NOT
'<!-- /*------------------------------*/ //-->

On Error GoTo ErrNext

        Dim strMandatory As Boolean
        Dim rstMandatory As DAO.Recordset

        Set rstMandatory = CurrentDb.OpenRecordset("SELECT " & myTableBE & ".[Mandatory] " & vbCrLf & _
            "FROM " & myTableBE & " " & vbCrLf & _
            "WHERE (((" & myTableBE & ".ID)>" & strCVID & ") AND ((" & myTableBE & ".[Mandatory]) = True)) " & vbCrLf & _
            "GROUP BY " & myTableBE & ".[Mandatory];")

        strMandatory = rstMandatory![Mandatory]
        rstMandatory.Close
        Set rstMandatory = Nothing

        FormName.[Mandatory].Value = strMandatory

ErrNext:
    'FormName.[Mandatory].Value = False
    Resume Next

'<!-- /*------------------------------*/ //-->
' THIS WILL ENABLE THE BUTTONS THE USER IS ALLOWED TO USE
' BASED ON THE VERSION THEY HAVE AND THE UPDATE PUSHED
'<!-- /*------------------------------*/ //-->

        'BUTTON DISPLAY VALUE
        buttonCaption = Array("Update Now", Null, "Maybe Later")

        'BUTTON VISIBILITY
        buttonVisible = Array(False, True, False)
        buttonEnabled = Array(False, True, False)

        If FormName.[Version_Current] <> FormName.[Version_Number] Then
            If FormName.[Mandatory] = True Then
                buttonCaption = Array(Null, "Update", Null)
            Else
                buttonVisible = Array(True, False, True)
                buttonEnabled = Array(True, False, True)
            End If
        Else
            buttonCaption = Array(Null, "Close", Null)
        End If

'<!-- /*------------------------------*/ //-->
'THIS WILL CHANGE THE BACKGROUND COLOR OF THE FORM
'<!-- /*------------------------------*/ //-->

        If FormName.[Version_Current] <> FormName.[Version_Number] Then
            'IF LAST UPDATE, OR ANY OF THE UPDATES MISSED, WHERE IMPORTANT THE FORM WILL TURN RED
            If FormName.[Mandatory] = True Then
                FormName.Detail.BackColor = 2500501 'RED
                FormName.[msg_info] = "IMPORTANT! To continue using this application, you are are required to install the available updates."
            Else
                'IF LAST UPDATE, OR ALL OF THE UPDATES MISSED WHERE NOT IMPORTANT THE FORM WILL TURN YELLOW
                FormName.Detail.BackColor = 2274542 'YELLOW
                FormName.[msg_info] = "NEW UPDATES! The available updates are not required by your application, and can be applied at any time."
            End If
        Else
            'IF THE USER HAS THE LATEST UPDATE THE FORM WILL TURN GREEN
            FormName.Detail.BackColor = 3107669 'GREEN
            FormName.[msg_info] = "You are currently running the lastest version of this application!"
        End If

    End If

'<!-- /*------------------------------*/ //-->
'APPLY LABEL FORMAT VARIABLES
'<!-- /*------------------------------*/ //-->
    Dim l As Integer
    For l = 0 To UBound(labelNameArray)
        Set labelName = labelNameArray(l)
        With labelName
            labelName.Caption = labelCaption(l)

            labelName.Width = twipsPerInch * labelWidth(l)
            labelName.Height = twipsPerInch * labelHeight(l)
            labelName.Move twipsPerInch * labelMoveLeft(l), twipsPerInch * labelMoveTop(l)

            labelName.FontSize = 11
            labelName.FontName = "Arial"
            labelName.FontWeight = fontWeight_Bold
            labelName.ForeColor = RGB(255, 255, 255)
            labelName.BackColor = RGB(255, 255, 255)
            labelName.BorderColor = RGB(127, 127, 127)

            labelName.BackStyle = 0
            labelName.BorderStyle = 0

            labelName.Visible = labelVisible(l)
        End With
    Next l

'<!-- /*------------------------------*/ //-->
'APPLY TEXTBOX FORMAT VARIABLES
'<!-- /*------------------------------*/ //-->
    Dim t As Integer
    For t = 0 To UBound(textboxNameArray)
        Set textboxName = textboxNameArray(t)
        With textboxName
            If FormName.Mandatory.Visible = False Then
                textboxName.Value = textboxValue(t)
            End If

            textboxName.Width = twipsPerInch * textboxWidth(t)
            textboxName.Height = twipsPerInch * textboxHeight(t)
            textboxName.Move twipsPerInch * textboxMoveLeft(t), twipsPerInch * textboxMoveTop(t)

            textboxName.FontSize = 11
            textboxName.FontName = "Arial"
            textboxName.FontWeight = fontWeight_Normal
            textboxName.ForeColor = RGB(64, 64, 64)
            textboxName.BackColor = RGB(255, 255, 255)

            If FormName.Mandatory.Visible = False Then
                textboxName.BackColor = RGB(217, 217, 217)
            End If

            textboxName.BorderColor = RGB(166, 166, 166)

            textboxName.BackStyle = 1
            textboxName.BorderStyle = 1

            textboxName.Visible = textboxVisible(t)
            textboxName.Enabled = textboxEnabled(t)
            textboxName.Locked = textboxLocked(t)
        End With
    Next t

'<!-- /*------------------------------*/ //-->
'APPLY BUTTON FORMAT VARIABLES
'<!-- /*------------------------------*/ //-->
    Dim b As Integer
    For b = 0 To UBound(buttonNameArray)
        Set buttonName = buttonNameArray(b)
        With buttonName
            buttonName.Caption = buttonCaption(b)

            buttonName.Width = twipsPerInch * buttonWidth(b)
            buttonName.Height = twipsPerInch * buttonHeight(b)
            buttonName.Move twipsPerInch * buttonMoveLeft(b), twipsPerInch * buttonMoveTop(b)

            buttonName.FontSize = 11
            buttonName.FontName = "Arial"
            buttonName.FontWeight = fontWeight_Bold
            buttonName.ForeColor = RGB(64, 64, 64)
            buttonName.BackColor = RGB(217, 217, 217)
            buttonName.BorderColor = RGB(166, 166, 166)

            buttonName.Visible = buttonVisible(b)
            buttonName.Enabled = buttonEnabled(b)

            buttonName.Controls (buttonCaption(b))
            buttonName.OnClick = "=" & buttonName.Name & "_Command()"

            buttonName.CursorOnHover = acCursorOnHoverHyperlinkHand
        End With
    Next b

'<!-- /*------------------------------*/ //-->
'FOOTER MESSAGE FORMAT
'<!-- /*------------------------------*/ //-->
    FormName.[msg_info].Visible = True
    FormName.[msg_info].Enabled = False
    FormName.[msg_info].Locked = True
    FormName.[msg_info].TextAlign = 2

    If FormName.Mandatory.Visible = False Then
        FormName.Mandatory.Visible = True

            FormName.[Mandatory].Width = twipsPerInch * 0.1806
            FormName.[Mandatory].Height = twipsPerInch * 0.1667
            FormName.[Mandatory].Move twipsPerInch * 3.9167, twipsPerInch * 0.5833

        FormName.[Updates].BackColor = RGB(255, 255, 255)

        FormName.[msg_info].FontSize = 28
        FormName.[msg_info].ForeColor = RGB(140, 140, 140)
        FormName.[msg_info].BackColor = RGB(37, 64, 97)
        FormName.[msg_info].FontWeight = fontWeight_Normal
    Else
        FormName.Mandatory.Visible = False

        FormName.[msg_info].FontSize = 12
        FormName.[msg_info].ForeColor = RGB(255, 0, 0)
        FormName.[msg_info].BackColor = RGB(255, 255, 255)
        FormName.[msg_info].FontWeight = fontWeight_Bold
    End If

End With

End Function

Public Function Update_Command()

    DoCmd.SetWarnings No
    On Error Resume Next

        If Forms(myForm).Update.Visible = True And Forms(myForm).Update.Enabled = True Then

            If Forms(myForm).Caption = "Push Software Update(s)" And Forms(myForm).Update.Caption = "Push" Then
                Forms(myForm).UpdateClose.Enabled = True
                'UPDATE VERSION ON THE FRONTEND TABLE
                DoCmd.RunSQL "UPDATE " & myTableFE & " SET " & myTableFE & ".VersionNumber = [Forms]![" & myForm & "]![Version_Number];"
                Run_Update_Push '()
            End If

            If Forms(myForm).Caption = "Check for Updates" And Forms(myForm).Update.Caption = "Update Now" And Forms(myForm).[Version_Current] <> Forms(myForm).[Version_Number] Then
                Run_Update_Install '()
            End If

        End If

    DoCmd.SetWarnings Yes

End Function

Public Function UpdateClose_Command()

    DoCmd.SetWarnings No
    On Error Resume Next

        If Forms(myForm).UpdateClose.Visible = True And Forms(myForm).UpdateClose.Enabled = True Then

            If Forms(myForm).Caption = "Push Software Update(s)" And Forms(myForm).UpdateClose.Caption = "Notify" Then
                Run_Update_Notify '()
            End If

            If Forms(myForm).Caption = "Check for Updates" Then
                If Forms(myForm).UpdateClose.Caption = "Update" And Forms(myForm).[Version_Current] <> Forms(myForm).[Version_Number] Then
                    Run_Update_Install '()
                Else
                    Close_Command '()
                End If
            End If

        End If

    DoCmd.SetWarnings Yes

End Function

Public Function Close_Command()
    DoCmd.Close acForm, myForm, acSaveNo
End Function

Public Function Run_Update_Notify()
'APPEND VERSION NUMBER AND INFORMATION TO BACKEND TABLE

    DoCmd.RunSQL "INSERT INTO " & myTableBE & " ( Version_Number, Updates, DTG, Updated_By, Mandatory, Application_Name ) SELECT [Forms]![" & myForm & "]![Version_Number] AS Version_Number, [Forms]![" & myForm & "]![Updates] AS Updates, [Forms]![" & myForm & "]![DTG] AS DTG, fSystemUserName() AS Updated_By, [Forms]![" & myForm & "]![Mandatory] AS Mandatory, fApplicationName() AS Application_Name;"

    MsgBox "Update notification has been sent to the users!"

End Function

Public Function Run_Update_Push()

Dim fileText As String
Dim fileScript, fileName

fileText = "@echo Off" & vbCrLf & _
"DEL " & """" & fSharedrivePath & "\" & "" & fApplicationName & """" & vbCrLf & _
"xcopy " & """C:\Users\" & fSystemUserName & "\Desktop\" & fApplicationName & """" & " /Y " & """" & fSharedrivePath & """" & vbCrLf & _
"DEL " & """C:\Users\" & fSystemUserName & "\Update_Push.bat""" & vbCrLf & _
"Exit"

Set fileScript = CreateObject("Scripting.FileSystemObject")
Set fileName = fileScript.CreateTextFile("C:\Users\" & fSystemUserName & "\Update_Push.bat", True)
    fileName.WriteLine (fileText)
    fileName.Close

Set fileName = Nothing
Set fileScript = Nothing

    Dim PathCrnt As String
    PathCrnt = "C:\Users\" & fSystemUserName & "\"
    Call Shell(PathCrnt & "Update_Push.bat")

    MsgBox "The new updates have been made available to the users!"

End Function

Public Function Run_Update_Install()

Dim fileText As String
Dim fileScript, fileName

fileText = "@echo Off" & vbCrLf & _
"taskkill /f /im MSACCESS.EXE" & vbCrLf & _
"DEL " & """C:\Users\" & fSystemUserName & "\Desktop\" & fApplicationName & """" & vbCrLf & _
"xcopy " & """" & fSharedrivePath & "\" & "" & fApplicationName & """" & " /Y " & """C:\Users\" & fSystemUserName & "\Desktop""" & vbCrLf & _
"start MSACCESS.EXE " & """C:\Users\" & fSystemUserName & "\Desktop\" & fApplicationName & """" & vbCrLf & _
"DEL " & """C:\Users\" & fSystemUserName & "\Update_Install.bat""" & vbCrLf & _
"Exit"

Set fileScript = CreateObject("Scripting.FileSystemObject")
Set fileName = fileScript.CreateTextFile("C:\Users\" & fSystemUserName & "\Update_Install.bat", True)
    fileName.WriteLine (fileText)
    fileName.Close

Set fileName = Nothing
Set fileScript = Nothing

    Dim PathCrnt As String
    PathCrnt = "C:\Users\" & fSystemUserName & "\"
    Call Shell(PathCrnt & "Update_Install.bat")

End Function

Function fComputerName() As String

'DIMENSION VARIABLES
Dim lngX As Long
Dim lngSize As Long
Dim stTemp As String

    'MAX API CALL CHAR LEN TO RETURN
    lngSize = 16

    'API CALL PLACEHOLDER
    stTemp = String$(lngSize, 0)

    'VALUE RETURNED BY API CALL
    lngX = GetComputerName(stTemp, lngSize)

    'GET STRIPPED COMPUTER NAME
    If lngX <> 0 Then
        fComputerName = Left$(stTemp, lngSize)
    Else
        fComputerName = ""
    End If
        
End Function
    
Function fSystemUserName() As String

'DIMENSION VARIABLES
Dim lngX As Long 'API VALUE
Dim lngSize As Long
Dim stTemp As String

    'MAX API CALL CHAR LEN TO RETURN
    lngSize = 24

    'API CALL PLACEHOLDER
    stTemp = String$(lngSize, 0)

    'VALUE RETURNED BY API CALL
    lngX = GetUserName(stTemp, lngSize)

    'GET COMPUTER NAME, MINUS NULL-STRING
    If lngX <> 0 Then
        fSystemUserName = Left$(stTemp, lngSize - 1)
    Else
        fSystemUserName = ""
    End If
    
End Function

Function fApplicationPath()
    'DATABASE FILE PATH
    fApplicationPath = [CurrentProject].[Path]
End Function

Function fApplicationName()
    'DATABASE NAME
    fApplicationName = [CurrentProject].[Name]
End Function
