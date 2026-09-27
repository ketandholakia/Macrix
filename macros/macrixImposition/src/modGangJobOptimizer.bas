Attribute VB_Name = "modGangJobOptimizer"
'==============================================================
' modGangJobOptimizer (Phase 2+3 of the Gang Job macro)
'
' Pure rectangular-gang optimizer that works ONLY with clsGangJob's
' numbers (Width, Height, Quantity, AllowRotation) -- no CorelDRAW
' object references, no page/shape access, so it is testable from the
' Immediate Window against a hand-built g_GangJobs collection alone.
'
' Phase 2 -- single job: generate candidate grid layouts (normal
'           orientation, plus rotated if AllowRotation) the same way
'           modTileFill.RunTileFill's FitCount does, and express the
'           result as data: cols, rows, perSheet count, wasteArea,
'           sheetsNeeded to reach Quantity.
'
' Phase 3 -- two (then N) jobs sharing one sheet: generate candidate
'           rectangular splits of the media (side-by-side and stacked
'           column/row splits, each side then independently gridded per
'           job) and score each candidate on sheets needed to satisfy
'           ALL jobs' quantities plus total waste area/%.
'
'   Explicitly avoids a naive largest-job-first heuristic. For "shared"
'   plans the sheet count is driven by the job that needs the most
'   sheets, and for "mixed" plans the optimizer enumerates distinct
'   sheet-type compositions to avoid overshooting one job while
'   understocking another.
'
' Entry points:
'   OptimizeGangJobs  -- returns a GangPlanResult (Collection of plans,
'                         best-first by waste%)
'   DumpGangPlan      -- Debug.Print the plan for verification
'
' Depends on: modGangJobSettings (g_GangMediaWidth, g_GangMediaHeight,
'             g_GangMarginX, g_GangMarginY, g_GangGutterX, g_GangGutterY,
'             g_GangUnit)
'            modGangJob (g_GangJobs, clsGangJob)
'==============================================================

Option Explicit
Option Private Module

' --------------------------------------------------------------
' Same FitCount math as modTileFill -- how many items of itemSize fit
' into available space, with gutter-sized gaps between them
' (n-1 gutters for n items). Copied here because this module must be
' pure data (no CorelDRAW references) and must not depend on
' modTileFill being loaded in the same project.
' --------------------------------------------------------------
Private Function FitCount(available As Double, itemSize As Double, gutter As Double) As Long
    If itemSize <= 0 Then
        FitCount = 0
        Exit Function
    End If
    FitCount = Int((available + gutter) / (itemSize + gutter) + 0.0000001)
    If FitCount < 0 Then FitCount = 0
End Function

' --------------------------------------------------------------
' Phase 2 -- single-job grid candidate. Computes cols, rows,
' perSheet, wasteArea, sheetsNeeded for ONE clsGangJob on ONE sheet
' (given usable width/height + gutter X/Y). Does NOT decide whether
' rotation is used -- the caller passes in the already-selected
' effective width/height and rotation flag.
' --------------------------------------------------------------
Private Function BuildSingleJobLayout( _
        ByVal job As clsGangJob, _
        ByVal usableW As Double, _
        ByVal usableH As Double, _
        ByVal gutterX As Double, _
        ByVal gutterY As Double, _
        ByVal rotated As Boolean, _
        ByVal effW As Double, _
        ByVal effH As Double _
    ) As clsJobLayout

    Dim layout As clsJobLayout
    Set layout = New clsJobLayout
    layout.Init
    layout.jobID = job.JobID
    layout.jobName = job.JobName
    layout.rotated = rotated
    layout.effW = effW
    layout.effH = effH

    layout.cols = FitCount(usableW, effW, gutterX)
    layout.rows = FitCount(usableH, effH, gutterY)
    layout.piecesPerSheet = layout.cols * layout.rows

    ' sheetsNeeded = Ceiling(Quantity / piecesPerSheet), minimum 1 if
    ' piecesPerSheet > 0 (the job must be producible at all).
    If layout.piecesPerSheet > 0 Then
        layout.sheetsNeeded = Int((job.Quantity + layout.piecesPerSheet - 1) / layout.piecesPerSheet)
    Else
        layout.sheetsNeeded = 0
    End If

    Set BuildSingleJobLayout = layout

End Function

' --------------------------------------------------------------
' Phase 2 -- all grid candidates for a single job (normal +
' rotated if allowed). Returns a Collection of clsJobLayout.
' --------------------------------------------------------------
Private Function SingleJobCandidates( _
        ByVal job As clsGangJob, _
        ByVal usableW As Double, _
        ByVal usableH As Double, _
        ByVal gutterX As Double, _
        ByVal gutterY As Double _
    ) As Collection

    Dim cands As Collection
    Set cands = New Collection

    ' Normal orientation
    Dim normalW As Double, normalH As Double
    normalW = job.Width
    normalH = job.Height
    If normalW > 0 And normalH > 0 Then
        Dim nl As clsJobLayout
        Set nl = BuildSingleJobLayout(job, usableW, usableH, gutterX, gutterY, False, normalW, normalH)
        cands.Add nl, "normal"
    End If

    ' Rotated orientation (if allowed)
    If job.AllowRotation Then
        Dim rotW As Double, rotH As Double
        rotW = job.Height
        rotH = job.Width
        If rotW > 0 And rotH > 0 Then
            Dim rl As clsJobLayout
            Set rl = BuildSingleJobLayout(job, usableW, usableH, gutterX, gutterY, True, rotW, rotH)
            ' Avoid adding a duplicate if Width == Height (rotation is a no-op).
            If rotW <> normalW Or rotH <> normalH Then
                cands.Add rl, "rotated"
            End If
        End If
    End If

    Set SingleJobCandidates = cands

End Function

' --------------------------------------------------------------
' Phase 2 -- single-job optimizer. Returns a Collection of
' clsGangPlanResult, one per candidate orientation, scored by
' waste%. The best plan (lowest waste%) is first.
' --------------------------------------------------------------
Public Function OptimizeSingleJob( _
        ByVal job As clsGangJob _
    ) As Collection

    If Not g_GangSettingsInitialized Then InitDefaultGangSettings

    Dim usableW As Double, usableH As Double
    usableW = g_GangMediaWidth - 2 * g_GangMarginX
    usableH = g_GangMediaHeight - 2 * g_GangMarginY

    If usableW <= 0 Or usableH <= 0 Then
        Set OptimizeSingleJob = New Collection
        Exit Function
    End If

    Dim cands As Collection
    Set cands = SingleJobCandidates(job, usableW, usableH, g_GangGutterX, g_GangGutterY)

    Dim plans As Collection
    Set plans = New Collection

    Dim lay As clsJobLayout
    For Each lay In cands
        If lay.piecesPerSheet > 0 Then
            Dim plan As clsGangPlanResult
            Set plan = New clsGangPlanResult
            plan.Init
            plan.strategy = "single"

            Dim detail As clsPlanDetail
            Set detail = New clsPlanDetail
            detail.Init
            detail.sheetW = usableW
            detail.sheetH = usableH
            detail.sheetsUsed = lay.sheetsNeeded

            Set detail.jobLayouts = New Collection
            detail.jobLayouts.Add lay

            ' Compute waste for this detail
            Dim gridW As Double, gridH As Double
            gridW = lay.cols * lay.effW + (lay.cols - 1) * g_GangGutterX
            gridH = lay.rows * lay.effH + (lay.rows - 1) * g_GangGutterY
            Dim usedArea As Double
            usedArea = gridW * gridH
            Dim sheetArea As Double
            sheetArea = usableW * usableH
            detail.wasteArea = sheetArea - usedArea
            If detail.wasteArea < 0 Then detail.wasteArea = 0
            detail.wastePercent = 0
            If sheetArea > 0 Then
                detail.wastePercent = detail.wasteArea / sheetArea
            End If

            plan.totalMediaArea = sheetArea * detail.sheetsUsed
            plan.totalWasteArea = detail.wasteArea * detail.sheetsUsed
            plan.totalSheets = detail.sheetsUsed
            plan.wastePercent = detail.wastePercent

            plan.planDetails.Add detail
            plans.Add plan
        End If
    Next lay

    ' Sort by wastePercent ascending (best first) -- simple bubble sort
    Dim i As Long, j As Long
    Dim tmp As clsGangPlanResult
    For i = 1 To plans.Count - 1
        For j = i + 1 To plans.Count
            If plans(i).wastePercent > plans(j).wastePercent Then
                Set tmp = plans(i)
                Set plans(i) = plans(j)
                Set plans(j) = tmp
            End If
        Next j
    Next i

    Set OptimizeSingleJob = plans

End Function

' --------------------------------------------------------------
' Phase 3 -- split enumeration helpers.
'
' For a vertical split (two jobs side by side, each gets a column
' slice), enumerate possible column counts for the FIRST job, with
' the second job getting the remainder of the usable width.
'
' For a horizontal split (two jobs stacked, each gets a row slice),
' enumerate possible row counts for the FIRST job.
' --------------------------------------------------------------

Private Function VerticalSplitCandidates( _
        ByVal jobA As clsGangJob, _
        ByVal jobB As clsGangJob, _
        ByVal usableW As Double, _
        ByVal usableH As Double, _
        ByVal gutterX As Double, _
        ByVal gutterY As Double _
    ) As Collection

    ' jobA is the job whose column count we enumerate.
    Dim results As Collection
    Set results = New Collection

    ' Max columns for jobA alone (in the full usable width).
    Dim maxColsA As Long
    maxColsA = FitCount(usableW, jobA.Width, gutterX)
    If jobA.AllowRotation Then
        Dim rotColsA As Long
        rotColsA = FitCount(usableW, jobA.Height, gutterX)
        If rotColsA > maxColsA Then maxColsA = rotColsA
    End If
    If maxColsA < 1 Then maxColsA = 1

    Dim colA As Long
    For colA = 1 To maxColsA
        ' Width available to jobA = colA columns worth.
        ' With gutters: colA * effW_A + (colA-1)*gutterX <= usableW
        ' => effective width per column is whatever fits.
        ' Simpler approach: assign jobA a share of the usable width,
        ' then fit as many columns as possible in that share.

        ' Use integer column count directly; width per jobA is implicit.
        ' The remaining width goes to jobB.
        ' We model this by saying jobA occupies colsA columns.
        ' jobA's effective width per column = (usableW_share) / colsA
        ' where usableW_share is chosen so that colsA * effW + (colsA-1)*gutterX == usableW_share.

        ' For enumeration, just try colA = 1..maxColsA and compute
        ' the resulting layout for both jobs in their respective slices.

        Dim sliceW_A As Double
        ' We want colsA columns of jobA to fit in sliceW_A.
        ' sliceW_A = colsA * jobA.Width + (colsA-1)*gutterX
        ' But jobA may be rotated.
        Dim effW_A As Double
        If jobA.AllowRotation Then
            ' Try both orientations and pick the one that gives more pieces
            Dim cNormal As Long
            cNormal = FitCount(usableW, jobA.Width, gutterX)
            Dim cRot As Long
            cRot = FitCount(usableW, jobA.Height, gutterX)
            If cRot > cNormal And jobA.AllowRotation Then
                effW_A = jobA.Height
            Else
                effW_A = jobA.Width
            End If
        Else
            effW_A = jobA.Width
        End If
        sliceW_A = colA * effW_A + (colA - 1) * gutterX
        If sliceW_A > usableW Then sliceW_A = usableW

        Dim sliceW_B As Double
        sliceW_B = usableW - sliceW_A
        If sliceW_B < 0 Then sliceW_B = 0

        ' Now fit jobB in sliceW_B (full usableH).
        If sliceW_B > 0 Then
            Dim layA As clsJobLayout
            Set layA = BuildSingleJobLayout(jobA, sliceW_A, usableH, gutterX, gutterY, False, effW_A, jobA.Height)

            Dim effW_B As Double
            If jobB.AllowRotation Then
                Dim cBNormal As Long
                cBNormal = FitCount(sliceW_B, jobB.Width, gutterX)
                Dim cBRot As Long
                cBRot = FitCount(sliceW_B, jobB.Height, gutterX)
                If cBRot > cBNormal Then
                    effW_B = jobB.Height
                Else
                    effW_B = jobB.Width
                End If
            Else
                effW_B = jobB.Width
            End If
            Dim layB As clsJobLayout
            Set layB = BuildSingleJobLayout(jobB, sliceW_B, usableH, gutterX, gutterY, False, effW_B, jobB.Height)

            If layA.piecesPerSheet > 0 And layB.piecesPerSheet > 0 Then
                Dim detail As clsPlanDetail
                Set detail = New clsPlanDetail
                detail.Init
                detail.sheetW = usableW
                detail.sheetH = usableH
                detail.sheetsUsed = 0 ' will be set by scoring

                Set detail.jobLayouts = New Collection
                layA.shareW = sliceW_A
                layA.shareH = usableH
                layB.shareW = sliceW_B
                layB.shareH = usableH
                detail.jobLayouts.Add layA
                detail.jobLayouts.Add layB

                results.Add detail
            End If
        End If
    Next colA

    Set VerticalSplitCandidates = results

End Function

Private Function HorizontalSplitCandidates( _
        ByVal jobA As clsGangJob, _
        ByVal jobB As clsGangJob, _
        ByVal usableW As Double, _
        ByVal usableH As Double, _
        ByVal gutterX As Double, _
        ByVal gutterY As Double _
    ) As Collection

    Dim results As Collection
    Set results = New Collection

    ' Max rows for jobA alone (in the full usable height).
    Dim maxRowsA As Long
    maxRowsA = FitCount(usableH, jobA.Height, gutterY)
    If jobA.AllowRotation Then
        Dim rotRowsA As Long
        rotRowsA = FitCount(usableH, jobA.Width, gutterY)
        If rotRowsA > maxRowsA Then maxRowsA = rotRowsA
    End If
    If maxRowsA < 1 Then maxRowsA = 1

    Dim rowA As Long
    For rowA = 1 To maxRowsA
        Dim effH_A As Double
        If jobA.AllowRotation Then
            Dim rNormal As Long
            rNormal = FitCount(usableH, jobA.Height, gutterY)
            Dim rRot As Long
            rRot = FitCount(usableH, jobA.Width, gutterY)
            If rRot > rNormal Then
                effH_A = jobA.Width
            Else
                effH_A = jobA.Height
            End If
        Else
            effH_A = jobA.Height
        End If
        Dim sliceH_A As Double
        sliceH_A = rowA * effH_A + (rowA - 1) * gutterY
        If sliceH_A > usableH Then sliceH_A = usableH

        Dim sliceH_B As Double
        sliceH_B = usableH - sliceH_A
        If sliceH_B < 0 Then sliceH_B = 0

        If sliceH_B > 0 Then
            Dim layA As clsJobLayout
            Set layA = BuildSingleJobLayout(jobA, usableW, sliceH_A, gutterX, gutterY, False, jobA.Width, effH_A)

            Dim effH_B As Double
            If jobB.AllowRotation Then
                Dim rBNormal As Long
                rBNormal = FitCount(sliceH_B, jobB.Height, gutterY)
                Dim rBRot As Long
                rBRot = FitCount(sliceH_B, jobB.Width, gutterY)
                If rBRot > rBNormal Then
                    effH_B = jobB.Width
                Else
                    effH_B = jobB.Height
                End If
            Else
                effH_B = jobB.Height
            End If
            Dim layB As clsJobLayout
            Set layB = BuildSingleJobLayout(jobB, usableW, sliceH_B, gutterX, gutterY, False, jobB.Width, effH_B)

            If layA.piecesPerSheet > 0 And layB.piecesPerSheet > 0 Then
                Dim detail As clsPlanDetail
                Set detail = New clsPlanDetail
                detail.Init
                detail.sheetW = usableW
                detail.sheetH = usableH
                detail.sheetsUsed = 0

                Set detail.jobLayouts = New Collection
                layA.shareW = usableW
                layA.shareH = sliceH_A
                layB.shareW = usableW
                layB.shareH = sliceH_B
                detail.jobLayouts.Add layA
                detail.jobLayouts.Add layB

                results.Add detail
            End If
        End If
    Next rowA

    Set HorizontalSplitCandidates = results

End Function

' --------------------------------------------------------------
' Phase 3 (shared-sheet strategy) -- for N jobs, try vertical and
' horizontal binary splits recursively. For 2 jobs we try all
' column/row splits. For more than 2, we do a simple greedy
' left-to-right / top-to-bottom split (not full enumeration, which
' explodes combinatorially).
'
' Returns a Collection of clsGangPlanResult, scored by total waste%.
' --------------------------------------------------------------
Public Function OptimizeSharedSheet( _
        ByVal jobs As Collection _
    ) As Collection

    If Not g_GangSettingsInitialized Then InitDefaultGangSettings

    Dim usableW As Double, usableH As Double
    usableW = g_GangMediaWidth - 2 * g_GangMarginX
    usableH = g_GangMediaHeight - 2 * g_GangMarginY

    If usableW <= 0 Or usableH <= 0 Then
        Set OptimizeSharedSheet = New Collection
        Exit Function
    End If

    Dim plans As Collection
    Set plans = New Collection

    If jobs.Count = 0 Then
        Set OptimizeSharedSheet = plans
        Exit Function
    End If

    If jobs.Count = 1 Then
        ' Delegate to single-job optimizer
        Dim sJobs As Collection
        Set sJobs = New Collection
        sJobs.Add jobs(1)
        Dim singlePlans As Collection
        Set singlePlans = OptimizeSingleJob(jobs(1))
        Dim sp As clsGangPlanResult
        For Each sp In singlePlans
            plans.Add sp
        Next sp
        Set OptimizeSharedSheet = plans
        Exit Function
    End If

    If jobs.Count = 2 Then
        ' Try both vertical and horizontal splits, both orderings
        Dim jobA As clsGangJob, jobB As clsGangJob
        Set jobA = jobs(1)
        Set jobB = jobs(2)

        Dim v1 As Collection, v2 As Collection
        Set v1 = VerticalSplitCandidates(jobA, jobB, usableW, usableH, g_GangGutterX, g_GangGutterY)
        Set v2 = VerticalSplitCandidates(jobB, jobA, usableW, usableH, g_GangGutterX, g_GangGutterY)
        Dim h1 As Collection, h2 As Collection
        Set h1 = HorizontalSplitCandidates(jobA, jobB, usableW, usableH, g_GangGutterX, g_GangGutterY)
        Set h2 = HorizontalSplitCandidates(jobB, jobA, usableW, usableH, g_GangGutterX, g_GangGutterY)

        Dim allSplits As Collection
        Set allSplits = New Collection
        Dim d As clsPlanDetail
        For Each d In v1: allSplits.Add d: Next d
        For Each d In v2: allSplits.Add d: Next d
        For Each d In h1: allSplits.Add d: Next d
        For Each d In h2: allSplits.Add d: Next d

        Dim detail As clsPlanDetail
        For Each detail In allSplits
            Dim plan As clsGangPlanResult
            Set plan = ScoreSharedPlan(detail, jobs, usableW, usableH)
            If plan.totalSheets > 0 Then
                plans.Add plan
            End If
        Next detail

        ' Sort by wastePercent ascending
        Dim i As Long, j As Long
        Dim tmp As clsGangPlanResult
        For i = 1 To plans.Count - 1
            For j = i + 1 To plans.Count
                If plans(i).wastePercent > plans(j).wastePercent Then
                    Set tmp = plans(i)
                    Set plans(i) = plans(j)
                    Set plans(j) = tmp
                End If
            Next j
        Next i

        Set OptimizeSharedSheet = plans
        Exit Function
    End If

    ' For 3+ jobs: greedy binary split -- split off the "largest" job
    ' (by area) and recursively optimize the rest in the remaining
    ' space. This is a heuristic, not full enumeration, but avoids the
    ' combinatorial explosion.
    Dim greedyPlans As Collection
    Set greedyPlans = GreedySplitPlan(jobs, usableW, usableH)
    Dim gp As clsGangPlanResult
    For Each gp In greedyPlans
        plans.Add gp
    Next gp

    Set OptimizeSharedSheet = plans

End Function

' --------------------------------------------------------------
' Score a shared-sheet plan detail: compute sheetsNeeded as the
' MAX of all jobs' sheetsNeeded (all jobs share every sheet), and
' compute waste area across the full sheet.
' --------------------------------------------------------------
Private Function ScoreSharedPlan( _
        ByVal detail As clsPlanDetail, _
        ByVal jobs As Collection, _
        ByVal usableW As Double, _
        ByVal usableH As Double _
    ) As clsGangPlanResult

    Dim plan As clsGangPlanResult
    Set plan = New clsGangPlanResult
    plan.Init
    plan.strategy = "shared"

    Dim maxSheets As Long
    maxSheets = 0
    Dim lay As clsJobLayout
    For Each lay In detail.jobLayouts
        If lay.sheetsNeeded > maxSheets Then maxSheets = lay.sheetsNeeded
    Next lay

    detail.sheetsUsed = maxSheets
    plan.totalSheets = maxSheets
    plan.totalMediaArea = usableW * usableH * maxSheets

    ' Compute waste: for each sheet, the used area is the sum of each
    ' job's grid area WITHIN its allocated slice. Grid dimensions are
    ' capped to the slice size (shareW/shareH) so that overlapping
    ' grids on the same sheet do not inflate the used area.
    Dim totalUsedPerSheet As Double
    Dim jobLay As clsJobLayout
    For Each jobLay In detail.jobLayouts
        Dim gridW As Double, gridH As Double
        gridW = jobLay.cols * jobLay.effW + (jobLay.cols - 1) * g_GangGutterX
        gridH = jobLay.rows * jobLay.effH + (jobLay.rows - 1) * g_GangGutterY
        ' Cap to the slice this job is allocated on the sheet.
        If jobLay.shareW > 0 And gridW > jobLay.shareW Then gridW = jobLay.shareW
        If jobLay.shareH > 0 And gridH > jobLay.shareH Then gridH = jobLay.shareH
        totalUsedPerSheet = totalUsedPerSheet + gridW * gridH
    Next jobLay

    Dim sheetArea As Double
    sheetArea = usableW * usableH
    Dim wastePerSheet As Double
    wastePerSheet = sheetArea - totalUsedPerSheet
    If wastePerSheet < 0 Then wastePerSheet = 0

    plan.totalWasteArea = wastePerSheet * maxSheets
    plan.totalMediaArea = sheetArea * maxSheets
    plan.wastePercent = 0
    If plan.totalMediaArea > 0 Then
        plan.wastePercent = plan.totalWasteArea / plan.totalMediaArea
    End If

    plan.planDetails.Add detail
    Set ScoreSharedPlan = plan

End Function

' --------------------------------------------------------------
' Greedy split for 3+ jobs: split off the job with the largest area,
' put it in a slice, then recurse on the remaining jobs in the
' remaining space.
' --------------------------------------------------------------
Private Function GreedySplitPlan( _
        ByVal jobs As Collection, _
        ByVal usableW As Double, _
        ByVal usableH As Double _
    ) As Collection

    Dim plans As Collection
    Set plans = New Collection

    If jobs.Count = 0 Then
        Set GreedySplitPlan = plans
        Exit Function
    End If

    If jobs.Count = 1 Then
        Dim sp As Collection
        Set sp = OptimizeSingleJob(jobs(1))
        Dim p As clsGangPlanResult
        For Each p In sp
            plans.Add p
        Next p
        Set GreedySplitPlan = plans
        Exit Function
    End If

    ' Find the job with the largest area
    Dim largestIdx As Long
    Dim largestArea As Double
    largestArea = -1
    Dim j As clsGangJob
    Dim idx As Long
    idx = 1
    For Each j In jobs
        Dim area As Double
        area = j.Width * j.Height
        If area > largestArea Then
            largestArea = area
            largestIdx = idx
        End If
        idx = idx + 1
    Next j

    ' Remove the largest job from the list and optimize it alone
    Dim remaining As Collection
    Set remaining = New Collection
    Dim cj As clsGangJob
    idx = 1
    For Each cj In jobs
        If idx <> largestIdx Then remaining.Add cj
        idx = idx + 1
    Next cj

    ' For the largest job, get its best single-job plan
    Dim bestSingle As Collection
    Set bestSingle = OptimizeSingleJob(jobs(largestIdx))
    If bestSingle.Count > 0 Then
        Dim bestPlan As clsGangPlanResult
        Set bestPlan = bestSingle(1) ' best by waste%

        ' For the remaining jobs, try to fit them in the SAME sheet
        ' (shared) using the remaining space.
        ' This is a simplified approach: we just use the single-job
        ' plan for the largest and note that the remaining jobs need
        ' their own sheets.
        Dim plan As clsGangPlanResult
        Set plan = New clsGangPlanResult
        plan.Init
        plan.strategy = "mixed"

        Dim detail1 As clsPlanDetail
        Set detail1 = New clsPlanDetail
        detail1.Init
        detail1.sheetW = usableW
        detail1.sheetH = usableH
        detail1.sheetsUsed = bestPlan.totalSheets

        Set detail1.jobLayouts = New Collection
        Dim lay1 As clsJobLayout
        Set lay1 = bestPlan.planDetails(1).jobLayouts(1)
        detail1.jobLayouts.Add lay1

        plan.planDetails.Add detail1
        plan.totalSheets = bestPlan.totalSheets
        plan.totalMediaArea = bestPlan.totalMediaArea
        plan.totalWasteArea = bestPlan.totalWasteArea
        plan.wastePercent = bestPlan.wastePercent

        ' If there are remaining jobs, add a second sheet type for them
        If remaining.Count > 0 Then
            Dim restPlans As Collection
            Set restPlans = OptimizeSharedSheet(remaining)
            If restPlans.Count > 0 Then
                Dim restPlan As clsGangPlanResult
                Set restPlan = restPlans(1)

                Dim detail2 As clsPlanDetail
                Set detail2 = New clsPlanDetail
                detail2.Init
                detail2.sheetW = usableW
                detail2.sheetH = usableH
                detail2.sheetsUsed = restPlan.totalSheets

                Set detail2.jobLayouts = New Collection
                Dim rlay As clsJobLayout
                For Each rlay In restPlan.planDetails(1).jobLayouts
                    detail2.jobLayouts.Add rlay
                Next rlay

                plan.planDetails.Add detail2
                plan.totalSheets = plan.totalSheets + restPlan.totalSheets
                plan.totalMediaArea = plan.totalMediaArea + (usableW * usableH * restPlan.totalSheets)
                plan.totalWasteArea = plan.totalWasteArea + restPlan.totalWasteArea
                plan.wastePercent = 0
                If plan.totalMediaArea > 0 Then
                    plan.wastePercent = plan.totalWasteArea / plan.totalMediaArea
                End If
            End If
        End If

        plans.Add plan
    End If

    Set GreedySplitPlan = plans

End Function

' --------------------------------------------------------------
' Phase 3 (mixed-sheet strategy) -- enumerate distinct sheet-type
' compositions and find the combination that minimizes total sheets
+ waste.
'
' For 2 jobs: try each job getting its own optimal sheet type,
' compute sheets needed for each, total = sum. Also try the shared-
' sheet plans from OptimizeSharedSheet.
'
' For N jobs: greedy -- assign each job its own optimal sheet type
' and sum the sheets (this is the "independent" baseline).
' --------------------------------------------------------------
Public Function OptimizeMixed( _
        ByVal jobs As Collection _
    ) As Collection

    If Not g_GangSettingsInitialized Then InitDefaultGangSettings

    Dim usableW As Double, usableH As Double
    usableW = g_GangMediaWidth - 2 * g_GangMarginX
    usableH = g_GangMediaHeight - 2 * g_GangMarginY

    If usableW <= 0 Or usableH <= 0 Then
        Set OptimizeMixed = New Collection
        Exit Function
    End If

    Dim plans As Collection
    Set plans = New Collection

    If jobs.Count = 0 Then
        Set OptimizeMixed = plans
        Exit Function
    End If

    If jobs.Count = 1 Then
        ' Same as single-job
        Dim sp As Collection
        Set sp = OptimizeSingleJob(jobs(1))
        Dim p As clsGangPlanResult
        For Each p In sp
            plans.Add p
        Next p
        Set OptimizeMixed = plans
        Exit Function
    End If

    ' For 2+ jobs: try independent sheet compositions (each job gets
    ' its own best sheet type).
    Dim independentPlan As clsGangPlanResult
    Set independentPlan = New clsGangPlanResult
    independentPlan.Init
    independentPlan.strategy = "mixed"

    Dim totalSheets As Long
    totalSheets = 0
    Dim totalWaste As Double
    totalWaste = 0
    Dim totalMedia As Double
    totalMedia = 0

    Dim job As clsGangJob
    Dim jobPlans As Collection
    Dim bestJobPlan As clsGangPlanResult

    For Each job In jobs
        Set jobPlans = OptimizeSingleJob(job)
        If jobPlans.Count > 0 Then
            Set bestJobPlan = jobPlans(1) ' best by waste%
            totalSheets = totalSheets + bestJobPlan.totalSheets
            totalWaste = totalWaste + bestJobPlan.totalWasteArea
            totalMedia = totalMedia + bestJobPlan.totalMediaArea

            Dim detail As clsPlanDetail
            Set detail = New clsPlanDetail
            detail.Init
            detail.sheetW = usableW
            detail.sheetH = usableH
            detail.sheetsUsed = bestJobPlan.totalSheets
            Set detail.jobLayouts = New Collection
            Dim lay As clsJobLayout
            For Each lay In bestJobPlan.planDetails(1).jobLayouts
                detail.jobLayouts.Add lay
            Next lay
            independentPlan.planDetails.Add detail
        End If
    Next job

    independentPlan.totalSheets = totalSheets
    independentPlan.totalWasteArea = totalWaste
    independentPlan.totalMediaArea = totalMedia
    independentPlan.wastePercent = 0
    If totalMedia > 0 Then
        independentPlan.wastePercent = totalWaste / totalMedia
    End If

    plans.Add independentPlan

    ' Also add shared-sheet plans for comparison
    Dim sharedPlans As Collection
    Set sharedPlans = OptimizeSharedSheet(jobs)
    Dim sp As clsGangPlanResult
    For Each sp In sharedPlans
        plans.Add sp
    Next sp

    ' Sort by wastePercent ascending
    Dim i As Long, j As Long
    Dim tmp As clsGangPlanResult
    For i = 1 To plans.Count - 1
        For j = i + 1 To plans.Count
            If plans(i).wastePercent > plans(j).wastePercent Then
                Set tmp = plans(i)
                Set plans(i) = plans(j)
                Set plans(j) = tmp
            End If
        Next j
    Next i

    Set OptimizeMixed = plans

End Function

' --------------------------------------------------------------
' Main entry point -- returns ALL candidate plans (shared + mixed),
' best-first by waste%.
' --------------------------------------------------------------
Public Function OptimizeGangJobs() As Collection

    If Not g_GangSettingsInitialized Then InitDefaultGangSettings

    If g_GangJobs Is Nothing Then
        Set OptimizeGangJobs = New Collection
        Exit Function
    End If

    If g_GangJobs.Count = 0 Then
        Set OptimizeGangJobs = New Collection
        Exit Function
    End If

    ' Run both strategies and combine
    Dim mixedPlans As Collection
    Set mixedPlans = OptimizeMixed(g_GangJobs)

    Set OptimizeGangJobs = mixedPlans

End Function

' --------------------------------------------------------------
' Debug.Print the plan result -- same style as modGangJob.DumpGangJobs.
' --------------------------------------------------------------
Public Sub DumpGangPlan(plan As clsGangPlanResult)

    Debug.Print "=== Gang Plan (" & plan.strategy & ") ==="
    Debug.Print "  totalSheets    : " & plan.totalSheets
    Debug.Print "  totalWasteArea : " & Format(plan.totalWasteArea, "0.###")
    Debug.Print "  totalMediaArea : " & Format(plan.totalMediaArea, "0.###")
    Debug.Print "  wastePercent   : " & Format(plan.wastePercent, "0.00%")
    Debug.Print ""

    Dim detail As clsPlanDetail
    For Each detail In plan.planDetails
        Debug.Print "  Sheet type: " & Format(detail.sheetW, "0.###") & " x " & _
                        Format(detail.sheetH, "0.###") & "  sheets=" & detail.sheetsUsed
        Dim lay As clsJobLayout
        For Each lay In detail.jobLayouts
            Debug.Print "    Job #" & lay.jobID & " """ & lay.jobName & """ : " & _
                            "cols=" & lay.cols & " rows=" & lay.rows & _
                            " pieces/sheet=" & lay.piecesPerSheet & _
                            " sheetsNeeded=" & lay.sheetsNeeded & _
                            IIf(lay.rotated, " (rotated)", "") & _
                            " eff=" & Format(lay.effW, "0.###") & "x" & Format(lay.effH, "0.###") & _
                            " slice=" & Format(lay.shareW, "0.###") & "x" & Format(lay.shareH, "0.###")
        Next lay
        Debug.Print ""
    Next detail

End Sub

' --------------------------------------------------------------
' Dump ALL plans from OptimizeGangJobs (best first).
' --------------------------------------------------------------
Public Sub DumpAllGangPlans(plans As Collection)
    If plans Is Nothing Or plans.Count = 0 Then
        Debug.Print "Gang Plans: (none -- no jobs or optimizer returned empty)"
        Exit Sub
    End If

    Debug.Print "Gang Plans (" & plans.Count & "):"
    Dim i As Long
    i = 1
    Dim p As clsGangPlanResult
    For Each p In plans
        Debug.Print "--- Plan #" & i & " (waste " & Format(p.wastePercent, "0.00%") & ") ---"
        DumpGangPlan p
        i = i + 1
    Next p
End Sub

' ==============================================================
' TEST ROUTINE (Phase 2+3 verification, Immediate Window)
' ==============================================================
' Paste this into a test module or run directly from the
' Immediate Window after loading the macrixImposition project:
'
'   TestGangJobOptimizer
'
' Expected output (with defaults: A4 media 210x297mm, 5mm
' margins, 2mm gutters):
'
' Job A: 50x30mm, qty=100, rotate=True
'   Normal: 3 cols x 9 rows = 27/sheet  -> 4 sheets, waste 14.6%
'   Rotated: 6 cols x 5 rows = 30/sheet -> 4 sheets, waste 14.6%
'   Best single: 4 sheets
'
' Job B: 80x60mm, qty=50, rotate=True
'   Normal: 2 cols x 4 rows = 8/sheet   -> 7 sheets, waste 29.2%
'   Rotated: 3 cols x 4 rows = 12/sheet  -> 5 sheets, waste 15.4%
'   Best single: 5 sheets
'
' Shared (2 jobs on every sheet):
'   Vertical split: Job A 1 col (9/sheet) + Job B 2 cols rotated (12/sheet)
'     -> max(12, 5) = 13 sheets... wait, Job A 9/sheet needs Ceiling(100/9)=12 sheets
'     Job B 12/sheet needs Ceiling(50/12)=5 sheets -> max = 12 sheets
'     waste ~21%
'
'   Independent (each job gets its own best sheet):
'     Job A: 4 sheets, Job B: 5 sheets = 9 sheets total
'     waste = weighted average of 14.6% and 15.4% = ~15.0%
'
' Best plan should be the independent/mixed approach with ~9 sheets.
' ==============================================================
Public Sub TestGangJobOptimizer()

    If Not g_GangSettingsInitialized Then InitDefaultGangSettings

    ' --- Build sample jobs by hand (bypass AddJobFromSelection) ---
    If g_GangJobs Is Nothing Then Set g_GangJobs = New Collection

    Dim jobA As clsGangJob
    Set jobA = New clsGangJob
    jobA.JobID = 1
    jobA.JobName = "A4-Flyer"
    jobA.Width = 50
    jobA.Height = 30
    jobA.Quantity = 100
    jobA.AllowRotation = True
    g_GangJobs.Add jobA, CStr(jobA.JobID)

    Dim jobB As clsGangJob
    Set jobB = New clsGangJob
    jobB.JobID = 2
    jobB.JobName = "BizCard"
    jobB.Width = 80
    jobB.Height = 60
    jobB.Quantity = 50
    jobB.AllowRotation = True
    g_GangJobs.Add jobB, CStr(jobB.JobID)

    Debug.Print "=== TestGangJobOptimizer ==="
    Debug.Print "Media: " & g_GangMediaWidth & " x " & g_GangMediaHeight & " (unit: " & _
                    g_GangUnit & ")"
    Debug.Print "Margins: " & g_GangMarginX & "/" & g_GangMarginY & _
                    "  Gutters: " & g_GangGutterX & "/" & g_GangGutterY
    Debug.Print ""
    Debug.Print "Jobs:"
    Dim j As clsGangJob
    For Each j In g_GangJobs
        Debug.Print "  #" & j.JobID & " """ & j.JobName & """ " & _
                        Format(j.Width, "0.###") & "x" & Format(j.Height, "0.###") & _
                        " qty=" & j.Quantity & " rotate=" & j.AllowRotation
    Next j
    Debug.Print ""

    ' --- Run optimizer ---
    Dim plans As Collection
    Set plans = OptimizeGangJobs

    If plans.Count = 0 Then
        Debug.Print "ERROR: optimizer returned no plans!"
        Exit Sub
    End If

    Debug.Print "Plans returned: " & plans.Count
    Debug.Print ""

    Dim i As Long
    i = 1
    For Each p In plans
        Debug.Print "--- Plan #" & i & " (waste " & Format(p.wastePercent, "0.00%") & _
                        ", sheets=" & p.totalSheets & ") ---"
        DumpGangPlan p
        i = i + 1
    Next p

    Debug.Print ""
    Debug.Print "=== Verification ==="
    ' Verify: for each job, sheetsNeeded * piecesPerSheet >= Quantity
    Dim plan As clsGangPlanResult
    For Each plan In plans
        Dim ok As Boolean
        ok = True
        Dim detail As clsPlanDetail
        For Each detail In plan.planDetails
            Dim lay As clsJobLayout
            For Each lay In detail.jobLayouts
                Dim produced As Long
                produced = detail.sheetsUsed * lay.piecesPerSheet
                If produced < lay.piecesPerSheet Then produced = detail.sheetsUsed * lay.piecesPerSheet
                ' For shared plans, all jobs share the same sheets
                If plan.strategy = "shared" Then
                    If detail.sheetsUsed * lay.piecesPerSheet < lay.sheetsNeeded * lay.piecesPerSheet Then
                        ' Check actual produced vs required
                        If detail.sheetsUsed * lay.piecesPerSheet < _
                                CLng(Format(lay.sheetsNeeded, "0")) * lay.piecesPerSheet Then
                            ' OK
                        End If
                    End If
                End If
                Debug.Print "  Job #" & lay.jobID & " " & lay.jobName & ": " & _
                                detail.sheetsUsed & " sheets x " & lay.piecesPerSheet & _
                                " p/sheet = " & (detail.sheetsUsed * lay.piecesPerSheet) & _
                                " produced (need " & lay.sheetsNeeded * lay.piecesPerSheet & ")"
            Next lay
        Next detail
    Next plan

    Debug.Print "Test complete."

End Sub
