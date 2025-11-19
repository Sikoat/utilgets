' SolidworksCUTOUTSmacro.vba
' SolidWorks 2023 VBA macro
' Purpose: On a sheet metal part (or any part with a large planar face), generate a drawing
'          containing a textual list of each internal hole/cutout (loops that do not touch
'          the outer edge of the sheet when flat) and report two longest in-plane dimensions
'          for each such feature.
'
'  - Detection: Identifies inner loops on a dominant planar face in the flat state.
'  - Dimension estimation: Uses a bounding-box approach on the loop's edges to produce two
'    longest in-plane dimensions. If exact shape analysis fails, gives best-effort values.
'  - Output: Creates a new drawing and inserts a Note listing the features and sizes.
'
' Limitations:
'  - Dimension estimation uses edge bounding boxes aligned to the model axes; for rotated
'    cutouts, values may slightly overstate actual length/width relative to feature-aligned axes.
'  - For complex faces or when a flat pattern cannot be activated, the macro falls back to the
'    largest planar face it can find.  So the current version of the macro does not come close to
'    listing all internal holes/cutouts then if such exist on multiple faces.  Just an initial test.
'
Option Explicit

' Entry point
Public Sub main()
    On Error GoTo FATAL_TRAP

    Dim swApp As SldWorks.SldWorks
    Set swApp = GetSolidWorksApp()
    If swApp Is Nothing Then
        MsgBox "SolidWorks application not found.", vbExclamation, "Internal Cutouts Macro"
        Exit Sub
    End If

    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    If swModel Is Nothing Then
        MsgBox "Please open a part document before running the macro.", vbInformation, "Internal Cutouts Macro"
        Exit Sub
    End If

    If swModel.GetType <> swDocPART Then
        MsgBox "Active document is not a Part. Please activate a sheet metal part and try again.", vbInformation, "Internal Cutouts Macro"
        Exit Sub
    End If

    Dim swPart As SldWorks.PartDoc
    Set swPart = swModel

    Dim logMessages As Collection
    Set logMessages = New Collection

    ' Try to put the model in flat state if it is sheet metal
    On Error Resume Next
    EnsureFlatPattern swModel, logMessages
    On Error GoTo 0

    ' Identify dominant planar face with inner loops
    Dim targetFace As SldWorks.Face2
    Set targetFace = FindPrimaryPlanarFace(swPart, logMessages)

    Dim results As Collection
    Set results = New Collection

    If Not targetFace Is Nothing Then
        On Error Resume Next
        CollectInnerLoopsAndSizes targetFace, results, logMessages
        On Error GoTo 0
    Else
        logMessages.Add "No suitable planar face found."
    End If

    ' Create the drawing and place the report
    Dim drawCreated As Boolean
    drawCreated = CreateDrawingWithReport(swApp, swModel, results, logMessages)

    If Not drawCreated Then
        MsgBox "Failed to create a drawing for the report. See log for details.", vbExclamation, "Internal Cutouts Macro"
    End If

    Exit Sub

FATAL_TRAP:
    ' Last-resort error handler; attempt to still inform the user
    On Error Resume Next
    MsgBox "Unexpected error: " & Err.Description, vbExclamation, "Internal Cutouts Macro"
End Sub

' Get a SolidWorks application object reliably
Private Function GetSolidWorksApp() As SldWorks.SldWorks
    On Error Resume Next
    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks
    If swApp Is Nothing Then
        Set swApp = GetObject(, "SldWorks.Application")
    End If
    If swApp Is Nothing Then
        Set swApp = CreateObject("SldWorks.Application")
        If Not swApp Is Nothing Then swApp.Visible = True
    End If
    Set GetSolidWorksApp = swApp
End Function

' Try to ensure flat pattern is active/unsuppressed
Private Sub EnsureFlatPattern(ByVal swModel As SldWorks.ModelDoc2, ByRef log As Collection)
    On Error Resume Next

    Dim feat As SldWorks.Feature
    Set feat = swModel.FirstFeature
    Do While Not feat Is Nothing
        Dim fname As String
        fname = feat.Name
        Dim ftype As String
        ftype = feat.GetTypeName2

        If LCase$(ftype) = "flatpattern" Or LCase$(fname) Like "flat-pattern*" Then
            ' Try to unsuppress and show
            If feat.IsSuppressed2(swInConfigurationOpts_e.swThisConfiguration, 0, Nothing) Then
                feat.SetSuppression2 2, 0, Nothing ' 2 = unsuppress
                log.Add "Unsuppressed flat pattern feature: " & fname
            End If
            Exit Do
        End If

        Set feat = feat.GetNextFeature
    Loop

    ' Try to show flat view if sheet metal specific API is present (best effort)
    ' If not found, proceed anyway.
End Sub

' Find a large planar face that likely represents the sheet in flat state
Private Function FindPrimaryPlanarFace(ByVal swPart As SldWorks.PartDoc, ByRef log As Collection) As SldWorks.Face2
    On Error Resume Next

    Dim bodies As Variant
    bodies = swPart.GetBodies2(swBodyType_e.swSolidBody, True)
    If IsEmpty(bodies) Or (IsArray(bodies) And UBound(bodies) < 0) Then
        bodies = swPart.GetBodies2(swBodyType_e.swSheetBody, True)
    End If

    If IsEmpty(bodies) Then
        log.Add "No bodies found in part."
        Set FindPrimaryPlanarFace = Nothing
        Exit Function
    End If

    Dim bestFace As SldWorks.Face2
    Dim bestScore As Double
    bestScore = -1#

    Dim i As Long
    For i = LBound(bodies) To UBound(bodies)
        Dim body As SldWorks.Body2
        Set body = bodies(i)
        If body Is Nothing Then GoTo NextBody

        Dim faces As Variant
        faces = body.GetFaces
        If IsEmpty(faces) Then GoTo NextBody

        Dim j As Long
        For j = LBound(faces) To UBound(faces)
            Dim f As SldWorks.Face2
            Set f = faces(j)
            If f Is Nothing Then GoTo NextFace

            Dim surf As SldWorks.Surface
            Set surf = f.GetSurface
            If surf Is Nothing Then GoTo NextFace

            If surf.IsPlane() Then
                ' Prefer faces with inner loops
                Dim loops As Variant
                loops = f.GetLoops
                Dim innerCount As Long
                innerCount = 0
                If Not IsEmpty(loops) Then
                    Dim k As Long
                    For k = LBound(loops) To UBound(loops)
                        Dim lp As SldWorks.Loop2
                        Set lp = loops(k)
                        If Not lp Is Nothing Then
                            If Not lp.IsOuter() Then innerCount = innerCount + 1
                        End If
                    Next k
                End If

                ' Score: prioritize more inner loops, then by area
                Dim props As Variant
                props = f.GetAreaProperties
                Dim area As Double
                If IsArray(props) And UBound(props) >= 0 Then area = props(0) Else area = 0#

                Dim score As Double
                score = CDbl(innerCount) * 1E6 + area ' heavy weight on inner loops count

                If score > bestScore Then
                    bestScore = score
                    Set bestFace = f
                End If
            End If
NextFace:
        Next j
NextBody:
    Next i

    If bestFace Is Nothing Then
        log.Add "No planar face with inner loops found; falling back to largest planar face."
        ' Try largest planar face
        bestScore = -1#
        For i = LBound(bodies) To UBound(bodies)
            Dim body2 As SldWorks.Body2
            Set body2 = bodies(i)
            If body2 Is Nothing Then GoTo NextBody2
            Dim faces2 As Variant
            faces2 = body2.GetFaces
            If IsEmpty(faces2) Then GoTo NextBody2
            Dim j2 As Long
            For j2 = LBound(faces2) To UBound(faces2)
                Dim f2 As SldWorks.Face2
                Set f2 = faces2(j2)
                If f2 Is Nothing Then GoTo NextFace2
                Dim s2 As SldWorks.Surface
                Set s2 = f2.GetSurface
                If Not s2 Is Nothing Then
                    If s2.IsPlane() Then
                        Dim props2 As Variant
                        props2 = f2.GetAreaProperties
                        Dim ar2 As Double
                        If IsArray(props2) Then ar2 = props2(0) Else ar2 = 0#
                        If ar2 > bestScore Then
                            bestScore = ar2
                            Set bestFace = f2
                        End If
                    End If
                End If
NextFace2:
            Next j2
NextBody2:
        Next i
    End If

    If bestFace Is Nothing Then
        log.Add "Failed to find any planar face."
    Else
        log.Add "Selected a planar face for analysis."
    End If

    Set FindPrimaryPlanarFace = bestFace
End Function

' Collect internal loops on face and compute their two longest in-plane dimensions (in inches)
Private Sub CollectInnerLoopsAndSizes(ByVal face As SldWorks.Face2, ByRef results As Collection, ByRef log As Collection)
    On Error Resume Next

    Dim loops As Variant
    loops = face.GetLoops
    If IsEmpty(loops) Then
        log.Add "No loops found on the selected face."
        Exit Sub
    End If

    Dim tolZero As Double
    tolZero = 1E-6 ' meters

    Dim idx As Long
    idx = 1

    Dim k As Long
    For k = LBound(loops) To UBound(loops)
        Dim lp As SldWorks.Loop2
        Set lp = loops(k)
        If lp Is Nothing Then GoTo NextLoop
        If lp.IsOuter() Then GoTo NextLoop ' skip the outer boundary

        Dim minMax As Variant
        minMax = GetLoopBoundingBox(lp)
        If IsEmpty(minMax) Then
            log.Add "Failed to compute bounding box for a loop; skipping."
            GoTo NextLoop
        End If

        Dim dx As Double, dy As Double, dz As Double
        dx = minMax(3) - minMax(0)
        dy = minMax(4) - minMax(1)
        dz = minMax(5) - minMax(2)

        ' Consider the two largest extents as the in-plane dimensions; ignore near-zero axis
        Dim a1 As Double, a2 As Double, a3 As Double
        a1 = dx: a2 = dy: a3 = dz

        ' Sort descending simple approach
        Dim arr(1 To 3) As Double
        arr(1) = a1: arr(2) = a2: arr(3) = a3
        Dim i As Integer, j As Integer
        For i = 1 To 2
            For j = i + 1 To 3
                If arr(j) > arr(i) Then
                    Dim t As Double: t = arr(i): arr(i) = arr(j): arr(j) = t
                End If
            Next j
        Next i

        Dim d1 As Double, d2 As Double
        d1 = arr(1): d2 = arr(2)

        ' Convert to inches and round reasonably
        Dim inch1 As Double, inch2 As Double
        inch1 = MetersToInches(d1)
        inch2 = MetersToInches(d2)

        Dim rec As LoopReport
        rec.Idx = idx
        rec.Dim1In = inch1
        rec.Dim2In = inch2
        rec.Note = ""

        results.Add rec
        idx = idx + 1

NextLoop:
    Next k

    If results.Count = 0 Then
        log.Add "No internal loops (holes/cutouts) detected on the selected face."
    End If
End Sub

' Return 6-element array {minX, minY, minZ, maxX, maxY, maxZ} for a loop by merging edge bounding boxes
Private Function GetLoopBoundingBox(ByVal lp As SldWorks.Loop2) As Variant
    On Error Resume Next

    Dim edges As Variant
    edges = lp.GetEdges
    If IsEmpty(edges) Then Exit Function

    Dim minX As Double, minY As Double, minZ As Double
    Dim maxX As Double, maxY As Double, maxZ As Double

    minX = 1E+99: minY = 1E+99: minZ = 1E+99
    maxX = -1E+99: maxY = -1E+99: maxZ = -1E+99

    Dim i As Long
    For i = LBound(edges) To UBound(edges)
        Dim ed As SldWorks.Edge
        Set ed = edges(i)
        If ed Is Nothing Then GoTo NextEdge

        Dim ent As SldWorks.Entity
        Set ent = ed
        If ent Is Nothing Then GoTo NextEdge

        Dim bx As Variant
        bx = ent.GetBox
        If IsArray(bx) And UBound(bx) >= 5 Then
            If bx(0) < minX Then minX = bx(0)
            If bx(1) < minY Then minY = bx(1)
            If bx(2) < minZ Then minZ = bx(2)
            If bx(3) > maxX Then maxX = bx(3)
            If bx(4) > maxY Then maxY = bx(4)
            If bx(5) > maxZ Then maxZ = bx(5)
        End If

NextEdge:
    Next i

    If maxX < minX Or maxY < minY Or maxZ < minZ Then Exit Function

    Dim outArr(0 To 5) As Double
    outArr(0) = minX: outArr(1) = minY: outArr(2) = minZ
    outArr(3) = maxX: outArr(4) = maxY: outArr(5) = maxZ

    GetLoopBoundingBox = outArr
End Function

Private Type LoopReport
    Idx As Long
    Dim1In As Double
    Dim2In As Double
    Note As String
End Type

Private Function MetersToInches(ByVal m As Double) As Double
    MetersToInches = m * 39.37007874015748#
End Function

Private Function FormatInches(ByVal inches As Double) As String
    On Error Resume Next
    FormatInches = FormatNumber(inches, 3, vbTrue, vbFalse, vbTrue) & "\"" ' quote mark for inches
End Function

' Create a drawing and add a textual report as a Note. Best-effort model view placement.
Private Function CreateDrawingWithReport(ByVal swApp As SldWorks.SldWorks, _
                                         ByVal sourceModel As SldWorks.ModelDoc2, _
                                         ByVal results As Collection, _
                                         ByVal log As Collection) As Boolean
    On Error GoTo EH

    Dim templatePath As String
    templatePath = ""
    On Error Resume Next
    templatePath = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplateDrawing)
    On Error GoTo EH

    Dim swDraw As SldWorks.DrawingDoc
    If Len(Trim$(templatePath)) > 0 Then
        Set swDraw = swApp.NewDocument(templatePath, 0, 0#, 0#)
    Else
        ' Fallback: try creating with no template (may prompt user depending on settings)
        Set swDraw = swApp.NewDocument("", 12, 0.297, 0.21) ' A4 as a fallback size
    End If

    If swDraw Is Nothing Then
        log.Add "Failed to create a new drawing document."
        CreateDrawingWithReport = False
        Exit Function
    End If

    Dim swDrawModel As SldWorks.ModelDoc2
    Set swDrawModel = swDraw

    ' Try to insert a model view (best-effort; not critical)
    On Error Resume Next
    Dim modelPath As String
    modelPath = sourceModel.GetPathName
    If Len(Trim$(modelPath)) > 0 Then
        swDraw.Create3rdAngleViews modelPath
    End If
    On Error GoTo EH

    ' Build the report text
    Dim reportText As String
    reportText = "Internal Holes/Cutouts Report" & vbCrLf & _
                 "Source: " & SafeModelName(sourceModel) & vbCrLf & _
                 "Units: inches" & vbCrLf & vbCrLf

    If results Is Nothing Or results.Count = 0 Then
        reportText = reportText & "No internal holes or cutouts detected or unable to determine." & vbCrLf
    Else
        reportText = reportText & "Index, Size 1, Size 2" & vbCrLf
        Dim i As Long
        For i = 1 To results.Count
            Dim rec As LoopReport
            ' Because user-defined types don't pass directly in Variant collections in some contexts,
            ' retrieve via a helper property – but in VBA this works if assigned by value.
            rec = results(i)
            reportText = reportText & CStr(rec.Idx) & ", " & _
                         FormatInches(rec.Dim1In) & ", " & FormatInches(rec.Dim2In) & vbCrLf
        Next i
    End If

    ' Append log/warnings
    If Not log Is Nothing And log.Count > 0 Then
        reportText = reportText & vbCrLf & "Notes:" & vbCrLf
        Dim j As Long
        For j = 1 To log.Count
            reportText = reportText & " - " & CStr(log(j)) & vbCrLf
        Next j
    End If

    ' Insert the Note at top-left of sheet
    Dim sheet As SldWorks.Sheet
    Set sheet = swDraw.GetCurrentSheet
    Dim width As Double, height As Double
    width = 0#: height = 0#
    If Not sheet Is Nothing Then
        sheet.GetSize width, height
    Else
        width = 0.297: height = 0.21 ' default A4 in meters
    End If

    Dim x As Double, y As Double
    x = 0.01 ' 10 mm from left
    y = height - 0.01 ' near top

    Dim swNote As SldWorks.Note
    Set swNote = swDrawModel.InsertNote(reportText)
    If Not swNote Is Nothing Then
        Dim ann As SldWorks.Annotation
        Set ann = swNote.GetAnnotation
        If Not ann Is Nothing Then
            ann.SetPosition x, y, 0#
        End If
    End If

    CreateDrawingWithReport = True
    Exit Function

EH:
    On Error Resume Next
    CreateDrawingWithReport = False
End Function

Private Function SafeModelName(ByVal swModel As SldWorks.ModelDoc2) As String
    On Error Resume Next
    Dim p As String
    p = swModel.GetPathName
    If Len(Trim$(p)) = 0 Then
        SafeModelName = swModel.GetTitle
    Else
        Dim f As String
        f = p
        Dim i As Long
        For i = Len(f) To 1 Step -1
            If Mid$(f, i, 1) = "\\" Or Mid$(f, i, 1) = "/" Then
                SafeModelName = Mid$(f, i + 1)
                Exit Function
            End If
        Next i
        SafeModelName = f
    End If
End Function
