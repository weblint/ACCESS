Attribute VB_Name = "mdlTools"
Option Compare Database
Option Explicit
    
Private Type TOpenFileName
    nStructSize     As Long
    hwndOwner       As Long
    hInstance       As Long
    sFilter         As String
    sCustomFilter   As String
    nCustFilterSize As Long
    nFilterIndex    As Long
    sFile           As String
    nFileSize       As Long
    sFileTitle      As String
    nTitleSize      As Long
    sInitDir        As String
    sDlgTitle       As String
    Flags           As Long
    nFileOffset     As Integer
    nFileExt        As Integer
    sDefFileExt     As String
    nCustData       As Long
    fnHook          As Long
    sTemplateName   As String
End Type

Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias _
    "GetSaveFileNameA" (pOpenfilename As TOpenFileName) As Long
    
Const wdFormatXMLDocument = 12
Const wdFormatXMLDocumentMacroEnabled = 13
Const wdFormatDocument = 0

Public Function Kopffarbe() As Long
    Dim lngColor As Long
    lngColor = Nz(DLookup("Formularkopf_Farbe", "tblOptionen"), 0)
    Kopffarbe = lngColor
End Function

Public Function IstFormularGeoeffnet_(strFormular As String) As Boolean
    IstFormularGeoeffnet_ = SysCmd(acSysCmdGetObjectState, acForm, strFormular) > 0
End Function

Public Function Backendpfad() As String
    Dim strPfad As String
    strPfad = DLookup("Database", "MSysObjects", "Name = 'tbl_reports_local'")
    strPfad = Left(strPfad, InStrRev(strPfad, "\") - 1)
    Backendpfad = strPfad
End Function
    
Function OPENFILENAME(Optional StartDir As String, _
    Optional sTitle As String = "Datei auswählen:", _
    Optional sFilter As String = "Access-DB (*.accdb)|Alle Dateien (*.*)") As String
    Static sDir As String
    WizHook.Key = 51488399
    If Len(StartDir) = 0 Then
        If Len(sDir) = 0 Then
            StartDir = CurrentProject.path
        Else
            StartDir = sDir
        End If
    End If
    Call WizHook.GetFileName(Application.hWndAccessApp, _
        "Microsoft Access", sTitle, _
        "Öffnen", OPENFILENAME, _
        StartDir, sFilter, _
        0&, 0&, &H40, False)
    If Len(OPENFILENAME) > 0 Then
        sDir = Left(OPENFILENAME, InStrRev(OPENFILENAME, "\", , vbTextCompare))
    End If
End Function

Public Sub Test_TempTabelleErstellen()
    Dim strSource As String
    strSource = "qryAdressenSerienbrief"
    TempTabelleErstellen strSource
End Sub

Public Sub TempTabelleErstellen(strSource As String)
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim fldSource As DAO.Field
    Dim prp As DAO.Property
    
    On Error Resume Next
    dbc.Execute "DROP TABLE tblSerienbrief_Temp", dbFailOnError
    On Error GoTo 0
    Set tdf = dbc.CreateTableDef("tblSerienbrief_Temp")
  
    For Each fldSource In dbs.OpenRecordset("SELECT * FROM " & strSource & " WHERE 1 = 2").Fields
        Set fld = tdf.CreateField(fldSource.Name, dbText)
        If fldSource.Type = dbText Then
            fld.AllowZeroLength = True
        End If
        tdf.Fields.Append fld
    Next fldSource
    
    Set fld = tdf.CreateField("Serienbrief", dbBoolean)
    tdf.Fields.Append fld
    
    dbc.TableDefs.Append tdf
    
    Set fld = tdf.Fields("Serienbrief")
    Set prp = fld.CreateProperty("DisplayControl", DB_INTEGER, 106, True)
    On Error Resume Next
    fld.Properties.Delete "DisplayControl"
    On Error GoTo 0
    fld.Properties.Append prp
    
    dbc.Execute "INSERT INTO tblSerienbrief_Temp SELECT * FROM " & strSource & " IN '" & dbs.Name & "'", dbFailOnError
    dbc.Execute "UPDATE tblSerienbrief_Temp SET Serienbrief = TRUE", dbFailOnError
    
End Sub

Public Function GetSaveFile(Optional StartDir As String, Optional DefFileName As String, Optional sFilter As String = "Alle Dateien (*.*)", Optional Title As String) As String
    Dim uOFN As TOpenFileName
    Dim sExt As String
    If Len(StartDir) = 0 Then StartDir = CurrentProject.path
    With uOFN
        .nStructSize = Len(uOFN)
        .hwndOwner = Application.hWndAccessApp
        sFilter = sFilter & vbNullChar & vbNullChar
        .sFilter = sFilter
        .sInitDir = StartDir & vbNullChar
        .nFilterIndex = 0
        .sDlgTitle = Title
        .Flags = &H382000
        .sFile = Space$(256) & vbNullChar
        .nFileSize = Len(.sFile)
        If Len(DefFileName) <> 0 Then Mid(.sFile, 1) = DefFileName
        .sFileTitle = Space$(256) & vbNullChar
        .nTitleSize = Len(.sFileTitle)
        If GetSaveFileName(uOFN) Then
            GetSaveFile = Left(uOFN.sFile, InStr(.sFile, vbNullChar) - 1)
            If InStr(1, GetSaveFile, ".") = 0 Then
                sExt = Split(uOFN.sFilter, Chr$(0))(uOFN.nFilterIndex * 2 - 1)
                sExt = Mid(sExt, InStr(1, sExt, ".") + 1)
                GetSaveFile = GetSaveFile & "." & sExt
            End If
        Else
            GetSaveFile = ""
        End If
    End With
End Function

Public Function ISODatum(varDate As Variant)
    ISODatum = Format(varDate, "\#yyyy\/mm\/dd hh\:nn\:ss\#")
End Function

Public Sub SilentRequery(frm As Form, strPKField As String) ', ctlActiveControl As Control)
    Dim lngXErsterDatensatz As Long
    Dim lngDatensatz As Long
    Dim lngXMarkierterDatensatzVorRequery As Long
    Dim lngPositionMarkierterDatensatzVorRequery As Long
    Dim lngHoeheDetailbereich As Long
    Dim lngPositionErsterDatensatzVorRequery As Long
    If IsNull(frm(strPKField)) Then
        frm.Requery
        Exit Sub
    End If
    lngDatensatz = frm(strPKField)
    lngPositionMarkierterDatensatzVorRequery = frm.SelTop
    lngXMarkierterDatensatzVorRequery = frm.CurrentSectionTop
    frm.Requery
    If frm.DefaultView = 1 Then
        lngHoeheDetailbereich = frm.Section(0).Height
    ElseIf frm.DefaultView = 2 Then
        lngHoeheDetailbereich = frm.CurrentSectionTop
    End If
    lngXErsterDatensatz = frm.CurrentSectionTop
    frm.SelTop = frm.Recordset.RecordCount
    lngPositionErsterDatensatzVorRequery = (lngXMarkierterDatensatzVorRequery - lngXErsterDatensatz) / lngHoeheDetailbereich
    frm.SelTop = lngPositionMarkierterDatensatzVorRequery - lngPositionErsterDatensatzVorRequery
    frm.SelTop = lngPositionMarkierterDatensatzVorRequery
    frm.Painting = True
End Sub


