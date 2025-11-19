' SolidWorks 2023 VBA macro
'
Option Explicit

' ---------- USER SETTINGS (Edit to fit the situation) ----------
' Name of the reference plane to project NORMAL TO.
' If this plane does not exist, the macro will use the Current View as a fallback.
Const TARGET_PLANE_NAME As String = "Top Plane"   ' <-- EDIT HERE if needed

' Geometry simplification thresholds (micro-detail suppression)
Const MIN_FEATURE_SIZE_INCH As Double = 0.01       ' smallest detail scale cared about (inches)
Const TINY_FILLET_RADIUS_INCH As Double = 0.01     ' faces with fillet radius ≤ this are ignored (inches)

' Toggle filters
Const ENABLE_FILTER_TINY_FACES As Boolean = True   ' ignore faces smaller than ~MIN_FEATURE_SIZE_INCH^2
Const IGNORE_TINY_FILLET_FACES As Boolean = True   ' ignore cylindrical faces with radius ≤ TINY_FILLET_RADIUS_INCH

' DXF output options
Const OUTPUT_SPLINES_AS_POLYLINES As Boolean = True

' ---------------------------------------------------------------

' SolidWorks objects
Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim swPart As SldWorks.PartDoc

Sub main()
    On Error GoTo EH

    Set swApp = Application.SldWorks
    If swApp Is Nothing Then
        MsgBox "SolidWorks application not available.", vbCritical
        Exit Sub
    End If

    Set swModel = swApp.ActiveDoc
    If swModel Is Nothing Then
        MsgBox "Open a part (.sldprt) first and make it the active document.", vbExclamation
        Exit Sub
    End If
    If swModel.GetType <> swDocPART Then
        MsgBox "Active document is not a part. Open a .sldprt and try again.", vbExclamation
        Exit Sub
    End If
    Set swPart = swModel

    ' Ensure model is rebuilt before processing
    On Error Resume Next
    swModel.ForceRebuild3 False
    On Error GoTo EH

    ' Compute thresholds in API (SI) units
    Dim inchToMeter As Double: inchToMeter = 0.0254
    Dim tinyRadiusM As Double: tinyRadiusM = TINY_FILLET_RADIUS_INCH * inchToMeter
    Dim minFaceAreaM2 As Double
    minFaceAreaM2 = (MIN_FEATURE_SIZE_INCH * inchToMeter) * (MIN_FEATURE_SIZE_INCH * inchToMeter)

    ' Find the largest solid body (prefer visible, fallback to hidden)
    Dim vBodies As Variant
    vBodies = swPart.GetBodies2(swSolidBody, True)
    If IsEmpty(vBodies) Then
        ' Try all solid bodies (including hidden)
        vBodies = swPart.GetBodies2(swSolidBody, False)
        If IsEmpty(vBodies) Then
            MsgBox "No solid bodies found in the part (including hidden).", vbExclamation
            Exit Sub
        Else
            swApp.SendMsgToUser2("No visible solid bodies. Using hidden bodies for export.", swMbWarning, swMbOk)
        End If
    End If

    Dim swBody As SldWorks.Body2
    Set swBody = PickLargestBody(vBodies)
    If swBody Is Nothing Then
        MsgBox "Failed to determine a solid body to export.", vbExclamation
        Exit Sub
    End If

    ' Build a filtered list of faces to export
    Dim faces As Variant: faces = swBody.GetFaces
    If IsEmpty(faces) Then
        MsgBox "No faces on selected body.", vbExclamation
        Exit Sub
    End If

    Dim picked As New Collection
    Dim i As Long
    For i = LBound(faces) To UBound(faces)
        Dim f As SldWorks.Face2
        Set f = faces(i)
        If FacePassesFilters(f, minFaceAreaM2, tinyRadiusM) Then
            picked.Add f
        End If
    Next i

    ' If filtered away too much, fall back to all faces
    If picked.Count = 0 Then
        For i = LBound(faces) To UBound(faces)
            picked.Add faces(i)
        Next i
    End If

    ' Prepare DXF export data
    Dim swDXF As SldWorks.ExportDxfData
    Set swDXF = swApp.GetExportFileData(swExportDataFileType_e.swExportDxfData)
    If swDXF Is Nothing Then
        MsgBox "Could not create DXF export data object. Please ensure SolidWorks is properly installed and try again.", vbCritical
        Exit Sub
    End If
    On Error Resume Next
    swDXF.SetExportGeometry swDxfExportGeometry_e.swDxfExportGeometry_EntitiesOnly  ' Faces/Loops/Edges path
    If OUTPUT_SPLINES_AS_POLYLINES Then swDXF.SetSplineAsPolyline True
    On Error GoTo EH

    ' Try to use the user-named plane; else fall back to current view
    Dim swPlaneFeat As SldWorks.Feature
    Dim planeSelected As Boolean: planeSelected = False
    
    ' Clear any pre-existing selections first
    swModel.ClearSelection2 True
    
    Set swPlaneFeat = swModel.FeatureByName(TARGET_PLANE_NAME)

    If Not swPlaneFeat Is Nothing Then
        ' Mark select plane to define projection reference
        planeSelected = swPlaneFeat.Select2(False, 0)
        If planeSelected Then
            ' Tell exporter to use the selected sketch/face as projection reference
            swDXF.SetProjectionType swDxfProjectionType_e.swDxfProjectionType_SketchOrFace
            swDXF.SetSketchOrFaceSelection True
        Else
            ' Selection failed; use current view orientation instead
            swDXF.SetProjectionType swDxfProjectionType_e.swDxfProjectionType_CurrentView
        End If
    Else
        ' No plane found — use current view orientation
        swDXF.SetProjectionType swDxfProjectionType_e.swDxfProjectionType_CurrentView
    End If

    ' Now select all filtered faces for export (Faces/Loops/Edges)
    If Not planeSelected Then swModel.ClearSelection2 True
    Dim added As Long: added = 0
    For i = 1 To picked.Count
        Dim ok As Boolean
        ok = picked(i).Select4(True, Nothing)
        If ok Then added = added + 1
    Next i

    If added = 0 Then
        MsgBox "Failed to select faces for DXF export.", vbCritical
        Exit Sub
    End If

    ' Determine output path
    Dim savePath As String
    savePath = BuildOutputPath(swModel)

    ' Perform export
    Dim errs As Long, warns As Long
    Dim okExport As Boolean
    okExport = swModel.Extension.SaveAs(savePath, swSaveAsCurrentVersion, swSaveAsOptions_Silent, swDXF, Nothing, errs, warns)

    ' Clear selection
    swModel.ClearSelection2 True

    If okExport Then
        swApp.SendMsgToUser2("AutoFlat DXF created: " & savePath, swMbInformation, swMbOk)
    Else
        Dim msg As String
        msg = "DXF export failed. err=" & errs & ", warn=" & warns & vbCrLf & _
              "Try saving the part to disk first, or reduce filters in macro header."
        swApp.SendMsgToUser2(msg, swMbStop, swMbOk)
    End If

    Exit Sub

EH:
    On Error Resume Next
    swModel.ClearSelection2 True
    MsgBox "Unexpected error: " & Err.Description, vbExclamation
End Sub

Private Function PickLargestBody(vBodies As Variant) As SldWorks.Body2
    On Error GoTo Fallback
    Dim i As Long
    Dim best As SldWorks.Body2
    Dim bestVol As Double: bestVol = -1#
    For i = LBound(vBodies) To UBound(vBodies)
        Dim b As SldWorks.Body2
        Set b = vBodies(i)
        If Not b Is Nothing Then
            Dim vol As Double
            vol = b.GetVolume ' m^3
            If vol > bestVol Then
                Set best = b
                bestVol = vol
            End If
        End If
    Next i
    Set PickLargestBody = best
    Exit Function
Fallback:
    ' If volume API fails, fall back to most faces
    Dim j As Long, bestCount As Long: bestCount = -1
    For j = LBound(vBodies) To UBound(vBodies)
        Dim bb As SldWorks.Body2
        Set bb = vBodies(j)
        If Not bb Is Nothing Then
            Dim vf As Variant: vf = bb.GetFaces
            Dim cnt As Long: cnt = 0
            If Not IsEmpty(vf) Then cnt = (UBound(vf) - LBound(vf) + 1)
            If cnt > bestCount Then
                bestCount = cnt
                Set best = bb
            End If
        End If
    Next j
    Set PickLargestBody = best
End Function

Private Function FacePassesFilters(f As SldWorks.Face2, minAreaM2 As Double, tinyRadiusM As Double) As Boolean
    On Error GoTo SafeNo

    ' Default to True, then eliminate if it breaks filters
    FacePassesFilters = True

    If ENABLE_FILTER_TINY_FACES Then
        Dim a As Double
        a = f.GetArea
        If a > 0# And a < minAreaM2 Then
            FacePassesFilters = False
            Exit Function
        End If
    End If

    If IGNORE_TINY_FILLET_FACES Then
        Dim s As SldWorks.Surface
        Set s = f.GetSurface
        If Not s Is Nothing Then
            Dim isCyl As Boolean
            On Error Resume Next
            isCyl = s.IsCylinder
            On Error GoTo SafeNo
            If isCyl Then
                Dim p As Variant
                On Error Resume Next
                p = s.CylinderParams ' [0..2]=origin, [3..5]=axis, [6]=radius (m)
                On Error GoTo SafeNo
                If Not IsEmpty(p) Then
                    If UBound(p) >= 6 Then
                        Dim r As Double
                        r = p(6)
                        If r > 0# And r <= tinyRadiusM Then
                            FacePassesFilters = False
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    End If

    Exit Function

SafeNo:
    ' If anything goes wrong in interrogation, keep the face (robustness over-aggressiveness)
    FacePassesFilters = True
End Function

Private Function BuildOutputPath(m As SldWorks.ModelDoc2) As String
    Dim base As String
    Dim p As String: p = m.GetPathName
    Dim hasPath As Boolean: hasPath = (Len(p) > 0)
    If hasPath Then
        Dim extPos As Long: extPos = InStrRev(p, ".")
        Dim slashPos As Long: slashPos = InStrRev(p, "\")
        If extPos > 0 And extPos > slashPos Then
            base = Left$(p, extPos - 1)
        Else
            base = p
        End If
    Else
        ' Not saved yet; use current working directory or Desktop fallback
        Dim cwd As String
        cwd = Application.SldWorks.GetCurrentWorkingDirectory
        If Len(cwd) = 0 Then
            cwd = Environ$("USERPROFILE") & "\Desktop"
            If Len(Dir$(cwd, vbDirectory)) = 0 Then
                cwd = "C:\"
            End If
        End If
        If Right$(cwd, 1) <> "\" Then
            cwd = cwd & "\"
        End If
        base = cwd & "AutoFlat_" & Format(Now, "yyyymmdd_HHMMSS")
    End If
    BuildOutputPath = base & "_AutoFlat.dxf"
End Function
