Attribute VB_Name = "SolidworksINFOTESTmacro"
Option Explicit
'
' Purpose: When run on an assembly, automatically creates a drawing, places a model view,
'          inserts a Bill of Materials (BOM) if possible, and adds useful notes
'          (assembly name, date/time, mass, predominant sheet-metal thickness guess).
'
'  - Open an assembly in SolidWorks 2023.
'  - Run this macro (import .vba into a macro .swp or paste into the macro editor).
'
' Relies on at least one .drwdot existing on the system.
' -------------------- Constants (duplicated enums) --------------------
Private Const swDocPART As Long = 1
Private Const swDocASSEMBLY As Long = 2
Private Const swDocDRAWING As Long = 3

' Paper sizes (best-effort values for NewDrawing2; may vary by version). Not guaranteed used.
Private Const swDwgPaperA3 As Long = 12
Private Const swDwgPaperA2 As Long = 13
Private Const swDwgPaperA1 As Long = 14

' BOM types (best-effort). 0 = Top-level only, 1 = Parts only, 2 = Indented
Private Const swBomType_PartsOnly As Long = 1

' Table anchor corners (best-effort): 1 = Top-left, 2 = Top-right, 3 = Bottom-left, 4 = Bottom-right
Private Const swTableAnchor_TopLeft As Long = 1

' SaveAs options
Private Const swSaveAsOptions_Silent As Long = 1
Private Const swSaveAsOptions_Copy As Long = 2

' -------------------- Entry point --------------------
Public Sub main()
    On Error GoTo EH

    Dim swApp As Object
    Set swApp = GetSwApp()
    If swApp Is Nothing Then
        MsgBox "Unable to access SolidWorks application.", vbExclamation, "AutoBOMDrawing"
        Exit Sub
    End If

    Dim swModel As Object
    Set swModel = swApp.ActiveDoc
    If swModel Is Nothing Then
        MsgBox "No active document. Open an assembly and try again.", vbExclamation, "AutoBOMDrawing"
        Exit Sub
    End If

    Dim docType As Long
    docType = 0
    On Error Resume Next
    docType = swModel.GetType
    On Error GoTo 0
    If docType <> swDocASSEMBLY Then
        MsgBox "Active document is not an assembly. Please activate an assembly and run again.", vbInformation, "AutoBOMDrawing"
        Exit Sub
    End If

    ' Prepare drawing
    Dim drw As Object
    Set drw = CreateNewDrawing(swApp)
    If drw Is Nothing Then
        MsgBox "Failed to create a drawing. Ensure at least one drawing template (.drwdot) exists.", vbExclamation, "AutoBOMDrawing"
        Exit Sub
    End If

    ' Place a primary model view
    Dim assemPath As String
    assemPath = SafeEnsureModelPath(swModel)

    Dim v As Object
    Set v = PlacePrimaryView(drw, swModel, assemPath)

    ' Try to insert BOM (non-fatal if fails)
    Dim bomInserted As Boolean
    bomInserted = False
    If Not v Is Nothing Then
        bomInserted = InsertBOMTable(drw, v)
    End If

    ' Add useful notes
    AddUsefulNotes drw, swModel, bomInserted

    Exit Sub

EH:
    ' Generic trap to keep the macro resilient
    On Error Resume Next
    MsgBox "Unexpected error: " & Err.Description, vbExclamation, "AutoBOMDrawing"
End Sub

' -------------------- Core helpers --------------------
Private Function GetSwApp() As Object
    On Error Resume Next
    Dim swApp As Object
    Set swApp = Application.SldWorks
    If swApp Is Nothing Then Set swApp = GetObject(, "SldWorks.Application")
    If swApp Is Nothing Then Set swApp = CreateObject("SldWorks.Application")
    Set GetSwApp = swApp
End Function

Private Function CreateNewDrawing(swApp As Object) As Object
    On Error GoTo EH

    Dim drwTemplate As String
    drwTemplate = FindAnyDrawingTemplate()

    Dim drw As Object
    If Len(drwTemplate) > 0 Then
        Set drw = swApp.NewDocument(drwTemplate, 0, 0#, 0#)
    End If

    If drw Is Nothing Then
        ' Fallback: try creating a drawing by paper size (may not be available in all versions)
        On Error Resume Next
        Set drw = CallByName(swApp, "NewDrawing2", VbMethod, swDwgPaperA3)
        On Error GoTo EH
    End If

    If Not drw Is Nothing Then
        Set CreateNewDrawing = drw
        Exit Function
    End If

EH:
    Set CreateNewDrawing = Nothing
End Function

Private Function PlacePrimaryView(drw As Object, swModel As Object, assemPath As String) As Object
    On Error GoTo EH

    Dim drawView As Object
    Dim x As Double, y As Double
    x = 0.22: y = 0.15 ' meters; adjust as needed

    If Len(assemPath) > 0 Then
        ' Try standard model views in order of preference
        drawView = Nothing
        On Error Resume Next
        Set drawView = drw.CreateDrawViewFromModelView3(assemPath, "*Isometric", x, y, 0#)
        If drawView Is Nothing Then Set drawView = drw.CreateDrawViewFromModelView3(assemPath, "*Front", x, y, 0#)
        If drawView Is Nothing Then Set drawView = drw.CreateDrawViewFromModelView3(assemPath, "*Top", x, y, 0#)
        On Error GoTo EH
    End If

    If drawView Is Nothing Then
        ' Last resort: try to create standard 3rd angle views if possible
        On Error Resume Next
        drw.ActivateSheet "Sheet1"
        CallByName drw, "Create3rdAngleViews2", VbMethod, swModel
        ' Select one view as primary
        Dim v As Object
        Set v = CallByName(drw, "GetFirstView", VbMethod)
        If Not v Is Nothing Then Set drawView = CallByName(v, "GetNextView", VbMethod)
        On Error GoTo EH
    End If

    Set PlacePrimaryView = drawView
    Exit Function

EH:
    Set PlacePrimaryView = Nothing
End Function

Private Function InsertBOMTable(drw As Object, view As Object) As Boolean
    On Error GoTo EH

    Dim ok As Boolean: ok = False

    ' Ensure the view is selected; many BOM insertion methods require a view selection
    On Error Resume Next
    CallByName view, "Select", VbMethod, False
    On Error GoTo 0

    Dim x As Double, y As Double
    x = 0.02: y = 0.19 ' near top-left

    ' Try multiple known variants across SW versions
    On Error Resume Next
    ' Variant A: InsertBomTable4(anchor, x, y, bomType, configuration, number)
    CallByName drw, "InsertBomTable4", VbMethod, swTableAnchor_TopLeft, x, y, swBomType_PartsOnly, 0, 1
    ok = (Err.Number = 0)
    Err.Clear

    If Not ok Then
        ' Variant B: InsertBomTable2(view, anchor, x, y, bomType, number)
        CallByName drw, "InsertBomTable2", VbMethod, view, swTableAnchor_TopLeft, x, y, swBomType_PartsOnly, 1
        ok = (Err.Number = 0)
        Err.Clear
    End If

    If Not ok Then
        ' Variant C: InsertBomTable5(view, anchor, x, y, bomType, configOption, numbering, start)
        CallByName drw, "InsertBomTable5", VbMethod, view, swTableAnchor_TopLeft, x, y, swBomType_PartsOnly, 0, 1, 1
        ok = (Err.Number = 0)
        Err.Clear
    End If

    If Not ok Then
        ' Try via extension on selected view (rare)
        Dim ext As Object
        Set ext = drw.Extension
        If Not ext Is Nothing Then
            CallByName ext, "InsertBomTable2", VbMethod, swTableAnchor_TopLeft, x, y, swBomType_PartsOnly, 1
            ok = (Err.Number = 0)
            Err.Clear
        End If
    End If

    On Error Go To EH
    InsertBOMTable = ok
    Exit Function

EH:
    InsertBOMTable = False
End Function

Private Sub AddUsefulNotes(drw As Object, swModel As Object, bomInserted As Boolean)
    On Error Resume Next

    Dim sheet As Object
    Set sheet = CallByName(drw, "GetCurrentSheet", VbMethod)

    Dim sx As Double, sy As Double
    If Not sheet Is Nothing Then
        sx = CallByName(sheet, "GetWidth", VbMethod)
        sy = CallByName(sheet, "GetHeight", VbMethod)
    Else
        sx = 0.42: sy = 0.3
    End If

    Dim title As String
    title = SafeDocTitle(swModel)

    Dim massText As String
    massText = GetMassSummary(swModel)

    Dim thickText As String
    thickText = PredominantThicknessSummary(swModel)

    Dim bomText As String
    If bomInserted Then
        bomText = "BOM: Inserted"
    Else
        bomText = "BOM: Not inserted (skipped or API variant unavailable)"
    End If

    Dim txt As String
    txt = "Assembly: " & title & vbCrLf & _
          "Date: " & Format$(Now, "yyyy-mm-dd hh:nn") & vbCrLf & _
          massText & vbCrLf & _
          thickText & vbCrLf & _
          bomText

    Dim noteX As Double, noteY As Double
    noteX = 0.02: noteY = 0.02

    Dim ann As Object
    Set ann = drw.CreateText2(txt, noteX, noteY, 0#)
    If ann Is Nothing Then
        ' Fallback older API
        CallByName drw, "InsertNote", VbMethod, txt
    End If
End Sub

' -------------------- Utilities --------------------
Private Function FindAnyDrawingTemplate() As String
    On Error Resume Next

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim paths(1 To 6) As String
    Dim i As Long

    Dim programData As String
    programData = Environ$("ProgramData")
    If Len(programData) = 0 Then programData = "C:\\ProgramData"

    paths(1) = programData & "\SOLIDWORKS\SOLIDWORKS 2023\templates"
    paths(2) = programData & "\SolidWorks\SOLIDWORKS 2023\templates"
    paths(3) = programData & "\SOLIDWORKS\SOLIDWORKS 2024\templates" ' future-proof slight
    paths(4) = programData & "\SOLIDWORKS\SOLIDWORKS 2022\templates" ' nearby
    paths(5) = programData & "\SOLIDWORKS"
    paths(6) = programData

    ' First, non-recursive check of likely folders
    For i = 1 To 4
        If fso.FolderExists(paths(i)) Then
            Dim file As Object
            Dim folder As Object
            Set folder = fso.GetFolder(paths(i))
            For Each file In folder.Files
                If LCase$(fso.GetExtensionName(file.Path)) = "drwdot" Then
                    FindAnyDrawingTemplate = file.Path
                    Exit Function
                End If
            Next file
        End If
    Next i

    ' Recursive search in broader ProgramData\SOLIDWORKS
    For i = 5 To 6
        If fso.FolderExists(paths(i)) Then
            Dim found As String
            found = FindFileRecursive(paths(i), "drwdot", 4) ' limit depth for performance
            If Len(found) > 0 Then
                FindAnyDrawingTemplate = found
                Exit Function
            End If
        End If
    Next i

    FindAnyDrawingTemplate = ""
End Function

Private Function FindFileRecursive(rootPath As String, ext As String, maxDepth As Long) As String
    On Error Resume Next
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If maxDepth < 0 Then Exit Function

    Dim folder As Object
    Set folder = fso.GetFolder(rootPath)
    If folder Is Nothing Then Exit Function

    Dim file As Object
    For Each file In folder.Files
        If LCase$(fso.GetExtensionName(file.Path)) = LCase$(ext) Then
            FindFileRecursive = file.Path
            Exit Function
        End If
    Next file

    Dim subf As Object
    For Each subf In folder.SubFolders
        FindFileRecursive = FindFileRecursive(subf.Path, ext, maxDepth - 1)
        If Len(FindFileRecursive) > 0 Then Exit Function
    Next subf
End Function

Private Function SafeEnsureModelPath(swModel As Object) As String
    On Error Resume Next
    Dim p As String
    p = swModel.GetPathName
    If Len(p) > 0 Then
        SafeEnsureModelPath = p
        Exit Function
    End If

    ' Save a temporary copy in %TEMP% so a drawing view can be made from file path
    Dim tempDir As String
    tempDir = Environ$("TEMP")
    If Len(tempDir) = 0 Then tempDir = Environ$("TMP")
    If Len(tempDir) = 0 Then tempDir = "C:\\Temp"

    Dim nameGuess As String
    nameGuess = SafeDocTitle(swModel)
    If Len(nameGuess) = 0 Then nameGuess = "Assembly"

    Dim tempPath As String
    tempPath = tempDir & "\" & nameGuess & "_AutoBOM_Temp.SLDASM"

    Dim errs As Long, warns As Long
    errs = 0: warns = 0
    swModel.SaveAs3 tempPath, (swSaveAsOptions_Silent Or swSaveAsOptions_Copy), errs, warns

    If Dir$(tempPath) <> "" Then
        SafeEnsureModelPath = tempPath
    Else
        SafeEnsureModelPath = ""
    End If
End Function

Private Function SafeDocTitle(swModel As Object) As String
    On Error Resume Next
    Dim n As String
    n = swModel.GetTitle
    n = Replace(n, ".SLDASM", "", , , vbTextCompare)
    n = Replace(n, ".sldasm", "", , , vbTextCompare)
    n = Replace(n, ".SLDPRT", "", , , vbTextCompare)
    n = Replace(n, ".sldprt", "", , , vbTextCompare)
    SafeDocTitle = n
End Function

Private Function GetMassSummary(swModel As Object) As String
    On Error Resume Next

    Dim mass As Double
    mass = 0#

    ' Try various API paths
    Dim v As Variant
    v = Empty

    ' Try ModelDoc2::GetMassProperties
    v = swModel.GetMassProperties
    If IsArray(v) Then
        If UBound(v) >= 5 Then mass = CDbl(v(5))
    End If

    If mass <= 0# Then
        ' Try Extension.GetMassProperties2 if available
        Dim ext As Object
        Set ext = swModel.Extension
        If Not ext Is Nothing Then
            v = CallByName(ext, "GetMassProperties2", VbMethod, 1, 0)
            If IsArray(v) Then
                If UBound(v) >= 5 Then mass = CDbl(v(5))
            End If
        End If
    End If

    If mass > 0# Then
        GetMassSummary = "Mass (approx): " & FormatNumberSafe(mass, 3) & " kg"
    Else
        GetMassSummary = "Mass: unavailable"
    End If
End Function

Private Function PredominantThicknessSummary(swModel As Object) As String
    On Error GoTo EH

    Dim compStats As Object
    Set compStats = CreateObject("Scripting.Dictionary")

    Dim assem As Object
    Set assem = swModel

    ' Traverse lightweight/suppressed components carefully
    Dim swConf As Object
    Set swConf = CallByName(assem, "GetActiveConfiguration", VbMethod)
    If swConf Is Nothing Then GoTo Done

    Dim swRoot As Object
    Set swRoot = CallByName(swConf, "GetRootComponent3", VbMethod, True)
    If swRoot Is Nothing Then GoTo Done

    CollectComponentsThicknessStats swRoot, compStats

Done:
    Dim bestKey As String
    bestKey = ""
    Dim bestCount As Long
    bestCount = 0

    Dim k As Variant
    For Each k In compStats.Keys
        If compStats(k) > bestCount Then
            bestCount = compStats(k)
            bestKey = CStr(k)
        End If
    Next k

    If bestCount > 0 Then
        PredominantThicknessSummary = "Predominant sheet thickness: " & bestKey & " mm (by part count)"
    Else
        PredominantThicknessSummary = "Predominant sheet thickness: unavailable"
    End If
    Exit Function

EH:
    PredominantThicknessSummary = "Predominant sheet thickness: unavailable"
End Function

Private Sub CollectComponentsThicknessStats(swComp As Object, stats As Object)
    On Error Resume Next

    If swComp Is Nothing Then Exit Sub

    Dim isSupp As Boolean
    isSupp = False
    isSupp = CallByName(swComp, "IsSuppressed", VbMethod)
    If isSupp Then Exit Sub

    Dim mdl As Object
    Set mdl = CallByName(swComp, "GetModelDoc2", VbMethod)

    If Not mdl Is Nothing Then
        Dim tmm As Double
        tmm = TryGetPartSheetMetalThickness(mdl)
        If tmm > 0# Then
            Dim key As String
            key = FormatNumberSafe(tmm, 3)
            If Not stats.Exists(key) Then stats.Add key, 0
            stats(key) = stats(key) + 1
        End If
    End If

    ' Recurse into children
    Dim children As Variant
    children = CallByName(swComp, "GetChildren", VbMethod)
    If IsArray(children) Then
        Dim i As Long
        For i = LBound(children) To UBound(children)
            CollectComponentsThicknessStats children(i), stats
        Next i
    End If
End Sub

Private Function TryGetPartSheetMetalThickness(partModel As Object) As Double
    On Error GoTo EH

    ' Only for parts
    Dim t As Long
    t = 0
    t = partModel.GetType
    If t <> swDocPART Then Exit Function

    ' 1) Custom properties commonly used
    Dim thickness As Double
    thickness = 0#

    Dim mgr As Object
    Set mgr = partModel.Extension.CustomPropertyManager("")
    If Not mgr Is Nothing Then
        thickness = ReadThicknessProp(mgr, Array("Sheet Metal Thickness", "THICKNESS", "Thickness", "GAUGE", "Gauge"))
        If thickness > 0# Then TryGetPartSheetMetalThickness = thickness: Exit Function
    End If

    ' 2) Sheet metal feature data (if available)
    Dim feat As Object
    Set feat = partModel.FirstFeature
    Do While Not feat Is Nothing
        Dim tname As String
        tname = feat.GetTypeName2
        If LCase$(tname) = "sheetmetal" Then
            Dim def As Object
            Set def = feat.GetDefinition
            If Not def Is Nothing Then
                ' ISheetMetalFeatureData2.Thickness may exist; attempt both direct and via parameter
                On Error Resume Next
                Dim thk As Double
                thk = CallByName(def, "Thickness", VbGet)
                If thk <= 0# Then thk = CallByName(def, "Thickness", VbMethod)
                On Error GoTo EH
                If thk > 0# Then TryGetPartSheetMetalThickness = thk * 1000#: Exit Function ' meters->mm if API returns m
            End If
        End If
        Set feat = feat.GetNextFeature
    Loop

    ' 3) Try cut-list or evaluated custom properties
    If Not mgr Is Nothing Then
        thickness = ReadThicknessProp(mgr, Array("SW-Sheet Metal Thickness", "SheetMetalThickness", "SM-THICKNESS"))
        If thickness > 0# Then TryGetPartSheetMetalThickness = thickness: Exit Function
    End If

    Exit Function

EH:
    ' Ignore
End Function

Private Function ReadThicknessProp(mgr As Object, names As Variant) As Double
    On Error Resume Next
    Dim i As Long
    For i = LBound(names) To UBound(names)
        Dim valOut As String, resolved As String, wasResolved As Boolean
        valOut = "": resolved = "": wasResolved = False
        mgr.Get2 CStr(names(i)), valOut, resolved
        If Len(resolved) = 0 Then resolved = valOut
        Dim d As Double
        d = ParseNumberFromString(resolved)
        If d > 0# Then
            ' Heuristic: if looks like in meters (e.g., 0.003), convert to mm; if looks like inches, leave? assume mm if < 0.1
            If d < 0.1 Then d = d * 1000#
            ReadThicknessProp = d
            Exit Function
        End If
    Next i
End Function

Private Function ParseNumberFromString(s As String) As Double
    On Error Resume Next
    Dim i As Long, ch As String, buf As String
    buf = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "0" And ch <= "9") Or ch = "." Or ch = "," Or ch = "-" Then
            buf = buf & ch
        ElseIf Len(buf) > 0 Then
            Exit For
        End If
    Next i
    buf = Replace(buf, ",", ".")
    If Len(buf) = 0 Then
        ParseNumberFromString = 0#
    Else
        ParseNumberFromString = Val(buf)
    End If
End Function

Private Function FormatNumberSafe(ByVal d As Double, Optional ByVal decimals As Long = 3) As String
    On Error Resume Next
    FormatNumberSafe = CStr(Round(d, decimals))
End Function
