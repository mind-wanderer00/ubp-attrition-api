' ============================================================
' CREDIT CARD CUSTOMER ATTRITION — Predictive Modeling & Analysis
' Sections 1–5 + Production Vision (CONSOLIDATED, FINAL)
' ============================================================
' HOW TO USE:
' 1. Open PowerPoint -> Alt+F11 (VBA Editor)
' 2. DELETE any existing Module1 / Module2
' 3. Insert -> Module -> Paste this ENTIRE script
' 4. Press F5 or Run -> Run Sub -> BuildDeck
' ============================================================

Option Explicit

Private Const CLR_DARK As Long = &H3D2A1A
Private Const CLR_NAVY As Long = &H612716
Private Const CLR_CARD As Long = &H4A3A28
Private Const CLR_WHITE As Long = &HFFFFFF
Private Const CLR_ICE As Long = &HFCDCC8
Private Const CLR_ACCENT As Long = &H4882F0
Private Const CLR_RED As Long = &H4747E8
Private Const CLR_GREEN As Long = &H6DAA59
Private Const CLR_GRAY As Long = &H9E9E9E
Private Const CLR_LTGRAY As Long = &HD0D0D0
Private Const CLR_DARKTEXT As Long = &H4B4B4B
Private Const CLR_OFFWHITE As Long = &HFAF5F0
Private Const CLR_TEAL As Long = &H908020
Private Const CLR_GOLD As Long = &H39B8E8

Private Const FNT_HEAD As String = "Arial"
Private Const FNT_BODY As String = "Arial"

Private Const SW As Single = 13.333
Private Const SH As Single = 7.5

' ════════════════════════════════════════════════════════════
Sub BuildDeck()
    Dim prs As Presentation
    Set prs = Application.Presentations.Add(msoTrue)
    prs.PageSetup.SlideWidth = SW * 72
    prs.PageSetup.SlideHeight = SH * 72
    Do While prs.Slides.Count > 0: prs.Slides(1).Delete: Loop

    Call Slide_Title(prs)
    Call Slide_Agenda(prs)
    Call Slide_ProblemStatement(prs)
    Call Slide_BusinessValue(prs)
    Call Slide_DatasetOverview(prs)
    Call Slide_Section1_Divider(prs)
    Call Slide_ClassImbalance(prs)
    Call Slide_UnivariateContinuous(prs)
    Call Slide_DataQuality(prs)
    Call Slide_UnivariateCategorical(prs)
    Call Slide_BivariateAnalysis(prs)
    Call Slide_CorrelationMatrix(prs)
    Call Slide_Section1_Synthesis(prs)
    Call Slide_Section2_Divider(prs)
    Call Slide_Behavioral(prs)
    Call Slide_TenureLifecycle(prs)
    Call Slide_FinancialDemographic(prs)
    Call Slide_GeographicSocial(prs)
    Call Slide_Section2_Synthesis(prs)
    Call Slide_Section3_Divider(prs)
    Call Slide_ModelSetup(prs)
    Call Slide_DecisionTree(prs)
    Call Slide_XGBoost(prs)
    Call Slide_XGBoostFeatureImportance(prs)
    Call Slide_ModelSummary(prs)
    Call Slide_Section4_Divider(prs)
    Call Slide_PreprocessPipeline(prs)
    Call Slide_PreprocessSummary(prs)
    Call Slide_Section5_Divider(prs)
    Call Slide_ModelStrategy(prs)
    Call Slide_ModelResults(prs)
    Call Slide_FinalComparison(prs)
    Call Slide_FinalVerdict(prs)
    Call Slide_ProductionVision(prs)
    Call Slide_Section6_Divider(prs)
    Call Slide_APIEndpoints(prs)
    Call Slide_DriftMonitoring(prs)
    Call Slide_RecommenderEngine(prs)
    Call Slide_Closing(prs)

    MsgBox "Deck built: " & prs.Slides.Count & " slides.", vbInformation, "Done"
End Sub

' ════════════════════════════════════════════════════════════
' HELPERS
' ════════════════════════════════════════════════════════════
Private Sub SetDarkBg(sld As Slide)
    sld.FollowMasterBackground = msoFalse
    sld.Background.Fill.Solid
    sld.Background.Fill.ForeColor.RGB = CLR_DARK
End Sub

Private Sub SetLightBg(sld As Slide)
    sld.FollowMasterBackground = msoFalse
    sld.Background.Fill.Solid
    sld.Background.Fill.ForeColor.RGB = CLR_OFFWHITE
End Sub

Private Function AddBox(sld As Slide, L As Single, T As Single, W As Single, H As Single, FillClr As Long) As Shape
    Dim shp As Shape
    Set shp = sld.Shapes.AddShape(msoShapeRectangle, L * 72, T * 72, W * 72, H * 72)
    shp.Fill.Solid: shp.Fill.ForeColor.RGB = FillClr: shp.Line.Visible = msoFalse
    Set AddBox = shp
End Function

Private Function AddRoundBox(sld As Slide, L As Single, T As Single, W As Single, H As Single, FillClr As Long) As Shape
    Dim shp As Shape
    Set shp = sld.Shapes.AddShape(msoShapeRoundedRectangle, L * 72, T * 72, W * 72, H * 72)
    shp.Fill.Solid: shp.Fill.ForeColor.RGB = FillClr: shp.Line.Visible = msoFalse
    shp.Adjustments.Item(1) = 0.06
    Set AddRoundBox = shp
End Function

Private Function AddTxt(sld As Slide, L As Single, T As Single, W As Single, H As Single, _
    txt As String, sz As Single, clr As Long, Optional bld As Boolean = False, _
    Optional aln As Long = ppAlignLeft, Optional fnt As String = "") As Shape
    Dim shp As Shape
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, L * 72, T * 72, W * 72, H * 72)
    With shp.TextFrame2
        .TextRange.Text = txt
        .TextRange.Font.Size = sz
        .TextRange.Font.Fill.ForeColor.RGB = clr
        .TextRange.Font.Bold = IIf(bld, msoTrue, msoFalse)
        .TextRange.Font.Name = IIf(fnt = "", FNT_BODY, fnt)
        .TextRange.ParagraphFormat.Alignment = aln
        .WordWrap = msoTrue: .AutoSize = ppAutoSizeNone
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
    End With
    Set AddTxt = shp
End Function

Private Sub AddAccentBar(sld As Slide, L As Single, T As Single, H As Single, Optional clr As Long = -1)
    If clr = -1 Then clr = CLR_ACCENT
    Call AddBox(sld, L, T, 0.06, H, clr)
End Sub

Private Sub AddStat(sld As Slide, L As Single, T As Single, W As Single, _
    statVal As String, statLabel As String, Optional accentClr As Long = -1)
    If accentClr = -1 Then accentClr = CLR_ACCENT
    Call AddTxt(sld, L, T, W, 0.6, statVal, 36, accentClr, True, ppAlignCenter, FNT_HEAD)
    Call AddTxt(sld, L, T + 0.55, W, 0.4, statLabel, 12, CLR_GRAY, False, ppAlignCenter)
End Sub

Private Sub SlideTitle(sld As Slide, title As String, subtitle As String, Optional isDark As Boolean = True)
    Dim clrT As Long, clrS As Long
    If isDark Then clrT = CLR_WHITE: clrS = CLR_ICE Else clrT = CLR_DARK: clrS = CLR_GRAY
    Call AddTxt(sld, 0.8, 0.4, 11, 0.5, title, 28, clrT, True, ppAlignLeft, FNT_HEAD)
    If subtitle <> "" Then Call AddTxt(sld, 0.8, 0.95, 11, 0.35, subtitle, 14, clrS, False)
End Sub

Private Sub ContentSlideSetup(sld As Slide, title As String, subtitle As String)
    Call SetLightBg(sld)
    Call AddBox(sld, 0, 0, SW, 0.06, CLR_ACCENT)
    Call SlideTitle(sld, title, subtitle, False)
End Sub

Private Sub FormatTableCell(tbl As Table, r As Integer, c As Integer, txt As String, sz As Single, totalRows As Integer)
    tbl.Cell(r, c).Shape.TextFrame2.TextRange.Text = txt
    tbl.Cell(r, c).Shape.TextFrame2.TextRange.Font.Size = sz
    tbl.Cell(r, c).Shape.TextFrame2.TextRange.Font.Name = FNT_BODY
    tbl.Cell(r, c).Shape.TextFrame2.TextRange.ParagraphFormat.Alignment = ppAlignCenter
    If r = 1 Then
        tbl.Cell(r, c).Shape.Fill.ForeColor.RGB = CLR_DARK
        tbl.Cell(r, c).Shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = CLR_WHITE
        tbl.Cell(r, c).Shape.TextFrame2.TextRange.Font.Bold = msoTrue
    ElseIf r Mod 2 = 0 Then
        tbl.Cell(r, c).Shape.Fill.ForeColor.RGB = CLR_WHITE
        tbl.Cell(r, c).Shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = CLR_DARKTEXT
    Else
        tbl.Cell(r, c).Shape.Fill.ForeColor.RGB = CLR_OFFWHITE
        tbl.Cell(r, c).Shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = CLR_DARKTEXT
    End If
    tbl.Cell(r, c).Borders(ppBorderTop).ForeColor.RGB = CLR_LTGRAY
    tbl.Cell(r, c).Borders(ppBorderBottom).ForeColor.RGB = CLR_LTGRAY
    tbl.Cell(r, c).Borders(ppBorderLeft).ForeColor.RGB = CLR_LTGRAY
    tbl.Cell(r, c).Borders(ppBorderRight).ForeColor.RGB = CLR_LTGRAY
End Sub

Private Sub SectionDivider(prs As Presentation, secNum As String, secTitle As String, secDesc As String)
    Dim sld As Slide
    Set sld = prs.Slides.Add(prs.Slides.Count + 1, ppLayoutBlank)
    Call SetDarkBg(sld)
    Call AddTxt(sld, 0.8, 1.5, 5, 1.2, secNum, 96, CLR_ACCENT, True, ppAlignLeft, FNT_HEAD)
    Call AddBox(sld, 0.8, 3#, 3, 0.04, CLR_ACCENT)
    Call AddTxt(sld, 0.8, 3.3, 10, 0.8, secTitle, 36, CLR_WHITE, True, ppAlignLeft, FNT_HEAD)
    Call AddTxt(sld, 0.8, 4.2, 8, 0.8, secDesc, 18, CLR_ICE, False)
End Sub

' ════════════════════════════════════════════════════════════
' SLIDE 1: TITLE
' ════════════════════════════════════════════════════════════
Private Sub Slide_Title(prs As Presentation)
    Dim sld As Slide
    Set sld = prs.Slides.Add(prs.Slides.Count + 1, ppLayoutBlank)
    Call SetDarkBg(sld)
    Call AddBox(sld, 0, 0, 0.15, SH, CLR_ACCENT)
    Call AddBox(sld, 0.5, 2.6, 4.5, 0.01, CLR_ACCENT)
    Call AddBox(sld, 0.5, 5.2, 3, 0.01, CLR_ACCENT)
    Call AddTxt(sld, 0.8, 1.2, 10, 0.8, "CREDIT CARD", 48, CLR_WHITE, True, ppAlignLeft, FNT_HEAD)
    Call AddTxt(sld, 0.8, 1.85, 10, 0.8, "CUSTOMER ATTRITION", 48, CLR_ACCENT, True, ppAlignLeft, FNT_HEAD)
    Call AddTxt(sld, 0.8, 2.85, 8, 0.5, "Predictive Modeling & Analysis", 22, CLR_ICE, False)
    Call AddTxt(sld, 0.8, 3.4, 9, 0.4, "Sections 1" & ChrW(8211) & "6: Statistical Analysis  |  EDA  |  Signal Validation  |  Pre-processing  |  Modeling  |  Production", 14, CLR_GRAY, False)
    Dim statY As Single: statY = 4.3
    Call AddStat(sld, 1, statY, 2.5, "100K", "Customer Records")
    Call AddStat(sld, 3.8, statY, 2.5, "13", "Core Features")
    Call AddStat(sld, 6.6, statY, 2.5, "~5%", "Churn Rate")
    Call AddStat(sld, 9.4, statY, 2.5, "0.05", "Best PR-AUC")
    Call AddTxt(sld, 0.8, 6.8, 6, 0.3, "Confidential  |  Data Science Division", 12, CLR_GRAY, False)
End Sub

' ════════════════════════════════════════════════════════════
' SLIDE 2: AGENDA
' ════════════════════════════════════════════════════════════
Private Sub Slide_Agenda(prs As Presentation)
    Dim sld As Slide, i As Integer, yPos As Single
    Dim cardClr As Long, numClr As Long, txtClr As Long, badge As Shape
    Dim sections(1 To 6, 1 To 3) As String
    Set sld = prs.Slides.Add(prs.Slides.Count + 1, ppLayoutBlank)
    Call SetDarkBg(sld): Call AddBox(sld, 0, 0, 0.15, SH, CLR_ACCENT)
    Call AddTxt(sld, 0.8, 0.5, 10, 0.6, "DECK STRUCTURE", 32, CLR_WHITE, True, ppAlignLeft, FNT_HEAD)
    sections(1, 1) = "01": sections(1, 2) = "Statistical Analysis": sections(1, 3) = "Univariate, Bivariate, Multivariate, Significance Tests"
    sections(2, 1) = "02": sections(2, 2) = "Exploratory Data Analysis": sections(2, 3) = "Behavioral, Lifecycle, Demographic, Geographic deep-dive"
    sections(3, 1) = "03": sections(3, 2) = "Model-Based Signal Validation": sections(3, 3) = "Decision Tree + XGBoost as diagnostic tools"
    sections(4, 1) = "04": sections(4, 2) = "Pre-processing Pipeline": sections(4, 3) = "Three model-specific data tables, shared split"
    sections(5, 1) = "05": sections(5, 2) = "Model Development & Evaluation": sections(5, 3) = "LR, EBM, XGBoost — tuned and compared"
    sections(6, 1) = "06": sections(6, 2) = "Production Pipeline": sections(6, 3) = "FastAPI, PSI Drift Monitoring, Rule-Based Recommender, Docker"
    For i = 1 To 6
        yPos = 1.3 + (i - 1) * 0.95
        cardClr = CLR_CARD
        Call AddRoundBox(sld, 0.8, yPos, 11.5, 0.8, cardClr)
        numClr = CLR_ACCENT
        Call AddTxt(sld, 1.1, yPos + 0.1, 0.8, 0.55, sections(i, 1), 24, numClr, True, ppAlignLeft, FNT_HEAD)
        txtClr = CLR_WHITE
        Call AddTxt(sld, 2.1, yPos + 0.05, 5, 0.35, sections(i, 2), 16, txtClr, True)
        Call AddTxt(sld, 2.1, yPos + 0.42, 7, 0.3, sections(i, 3), 12, CLR_GRAY, False)
        Set badge = AddRoundBox(sld, 10.5, yPos + 0.2, 1.5, 0.38, CLR_GREEN)
        Call AddTxt(sld, 10.5, yPos + 0.22, 1.5, 0.33, "COMPLETE", 12, CLR_WHITE, True, ppAlignCenter)
    Next i
End Sub

' ════════════════════════════════════════════════════════════
' SLIDE 3: PROBLEM STATEMENT
' ════════════════════════════════════════════════════════════
Private Sub Slide_ProblemStatement(prs As Presentation)
    Dim sld As Slide, i As Integer, col As Integer, row As Integer, cx As Single, cy As Single
    Dim tasks(1 To 6, 1 To 2) As String
    Set sld = prs.Slides.Add(prs.Slides.Count + 1, ppLayoutBlank)
    Call ContentSlideSetup(sld, "Problem Statement & Background", "Business context driving this analysis")
    Call AddRoundBox(sld, 0.8, 1.6, 5.8, 2.2, CLR_WHITE)
    Call AddAccentBar(sld, 0.8, 1.6, 2.2)
    Call AddTxt(sld, 1.2, 1.75, 5, 0.35, "BUSINESS CHALLENGE", 15, CLR_ACCENT, True)
    Call AddTxt(sld, 1.2, 2.2, 5, 1.2, _
        "A major bank is experiencing customer churn in its credit card division. " & _
        "Understanding the drivers of attrition is critical for retention strategy " & _
        "design and improving customer satisfaction.", 14, CLR_DARKTEXT, False)
    Call AddRoundBox(sld, 7, 1.6, 5.5, 2.2, CLR_WHITE)
    Call AddAccentBar(sld, 7, 1.6, 2.2, CLR_TEAL)
    Call AddTxt(sld, 7.4, 1.75, 4.8, 0.35, "OBJECTIVE", 15, CLR_TEAL, True)
    Call AddTxt(sld, 7.4, 2.2, 4.8, 1.2, _
        "Build a predictive model to identify customers likely to close their " & _
        "credit card accounts, enabling proactive intervention before attrition occurs.", 14, CLR_DARKTEXT, False)
    Call AddTxt(sld, 0.8, 4.1, 5, 0.35, "TASK SCOPE", 16, CLR_DARK, True, ppAlignLeft, FNT_HEAD)
    tasks(1, 1) = "1. Data Understanding": tasks(1, 2) = "Missing values, outliers, duplicates, imbalance"
    tasks(2, 1) = "2. Feature Engineering": tasks(2, 2) = "Ratio features, log transforms, interactions"
    tasks(3, 1) = "3. Model Development": tasks(3, 2) = "LR, EBM, XGBoost, Random Forest"
    tasks(4, 1) = "4. Model Evaluation": tasks(4, 2) = "Accuracy, Precision, Recall, F1, ROC-AUC"
    tasks(5, 1) = "5. Insights & Strategy": tasks(5, 2) = "Key drivers, actionable bank strategies"
    tasks(6, 1) = "6. Dashboard (Bonus)": tasks(6, 2) = "Streamlit interactive dashboard"
    For i = 1 To 6
        col = ((i - 1) Mod 3): row = ((i - 1) \ 3)
        cx = 0.8 + col * 4.1: cy = 4.55 + row * 1.15
        Call AddRoundBox(sld, cx, cy, 3.8, 0.95, CLR_WHITE)
        Call AddTxt(sld, cx + 0.3, cy + 0.1, 3.3, 0.35, tasks(i, 1), 14, CLR_DARK, True)
        Call AddTxt(sld, cx + 0.3, cy + 0.48, 3.3, 0.4, tasks(i, 2), 12, CLR_GRAY, False)
    Next i
End Sub

' ════════════════════════════════════════════════════════════
' SLIDE 4: DATASET OVERVIEW
' ════════════════════════════════════════════════════════════
Private Sub Slide_DatasetOverview(prs As Presentation)
    Dim sld As Slide, i As Integer, cx As Single
    Dim cats(1 To 4, 1 To 2) As String, catClrs(1 To 4) As Long
    Dim cW As Single, gap As Single, startX As Single
    Set sld = prs.Slides.Add(prs.Slides.Count + 1, ppLayoutBlank)
    Call ContentSlideSetup(sld, "Dataset Overview", "101,000 synthetic customer records  |  13 core features  |  63-column full set")
    cats(1, 1) = "DEMOGRAPHIC": cats(1, 2) = "Age, Gender, MaritalStatus," & vbCrLf & "EducationLevel, Country"
    cats(2, 1) = "FINANCIAL": cats(2, 2) = "Income, CreditLimit," & vbCrLf & "TotalSpend"
    cats(3, 1) = "BEHAVIORAL": cats(3, 2) = "TotalTransactions," & vbCrLf & "Tenure, CardType"
    cats(4, 1) = "TARGET": cats(4, 2) = "AttritionFlag (binary:" & vbCrLf & "0 = Stayed, 1 = Attrited)"
    catClrs(1) = CLR_ACCENT: catClrs(2) = CLR_TEAL: catClrs(3) = CLR_NAVY: catClrs(4) = CLR_RED
    cW = 2.8: gap = 0.25: startX = (SW - (4 * cW + 3 * gap)) / 2
    For i = 1 To 4
        cx = startX + (i - 1) * (cW + gap)
        Call AddRoundBox(sld, cx, 1.7, cW, 2.2, CLR_WHITE)
        Call AddBox(sld, cx, 1.7, cW, 0.07, catClrs(i))
        Call AddTxt(sld, cx + 0.3, 1.95, cW - 0.6, 0.35, cats(i, 1), 14, catClrs(i), True)
        Call AddTxt(sld, cx + 0.3, 2.4, cW - 0.6, 1.2, cats(i, 2), 13, CLR_DARKTEXT, False)
    Next i
    Call AddTxt(sld, 0.8, 4.3, 5, 0.35, "KEY NUMBERS", 16, CLR_DARK, True, ppAlignLeft, FNT_HEAD)
    Call AddRoundBox(sld, 0.8, 4.85, 11.7, 1.6, CLR_WHITE)
    Call AddStat(sld, 1.3, 5, 2.2, "100,000", "Clean Records (post-dedup)", CLR_DARK)
    Call AddStat(sld, 3.8, 5, 2.2, "95 / 5", "Stayed vs. Attrited (%)", CLR_RED)
    Call AddStat(sld, 6.3, 5, 2.2, "13 + 50", "Core + Anonymous Features", CLR_TEAL)
    Call AddStat(sld, 8.8, 5, 2.8, "~5,000", "Missing per Column (~5%)", CLR_ACCENT)
End Sub

' ════════════════════════════════════════════════════════════
' SECTION DIVIDERS
' ════════════════════════════════════════════════════════════
Private Sub Slide_Section1_Divider(prs As Presentation)
    Call SectionDivider(prs, "01", "Statistical Analysis", "Univariate, Bivariate, Multivariate & Significance Tests")
End Sub
Private Sub Slide_Section2_Divider(prs As Presentation)
    Call SectionDivider(prs, "02", "EDA & Feature Discovery", "Answering four business questions about who churns and why")
End Sub
Private Sub Slide_Section3_Divider(prs As Presentation)
    Call SectionDivider(prs, "03", "Model-Based Signal Validation", "Decision Tree + XGBoost applied as diagnostic confirmation tools")
End Sub
Private Sub Slide_Section4_Divider(prs As Presentation)
    Call SectionDivider(prs, "04", "Pre-processing Pipeline", "Three model-specific data tables  |  Shared train/test indices  |  Fair comparison")
End Sub
Private Sub Slide_Section5_Divider(prs As Presentation)
    Call SectionDivider(prs, "05", "Model Development & Evaluation", "LR  |  EBM  |  XGBoost  " & ChrW(8212) & "  tuned, evaluated, compared on same test set")
End Sub

' ════════════════════════════════════════════════════════════
' SLIDE: CLASS IMBALANCE
' ════════════════════════════════════════════════════════════
Private Sub Slide_ClassImbalance(prs As Presentation)
    Dim sld As Slide
    Set sld = prs.Slides.Add(prs.Slides.Count + 1, ppLayoutBlank)
    Call ContentSlideSetup(sld, "1.1  Class Imbalance", "Target variable distribution " & ChrW(8212) & " a critical modeling consideration")
    Call AddRoundBox(sld, 0.8, 1.8, 9.5, 1.2, CLR_WHITE)
    Call AddBox(sld, 0.8, 1.8, 8.95, 1.2, CLR_TEAL)
    Call AddTxt(sld, 1.1, 1.92, 3, 0.4, "STAYED", 20, CLR_WHITE, True)
    Call AddTxt(sld, 1.1, 2.35, 5, 0.4, "95,040 customers  (95.04%)", 14, CLR_WHITE, False)
    Call AddRoundBox(sld, 0.8, 3.2, 9.5, 1.2, CLR_WHITE)
    Call AddBox(sld, 0.8, 3.2, 0.49, 1.2, CLR_RED)
    Call AddTxt(sld, 1.5, 3.32, 3, 0.4, "ATTRITED", 20, CLR_RED, True)
    Call AddTxt(sld, 1.5, 3.75, 5, 0.4, "4,960 customers  (4.96%)", 14, CLR_DARKTEXT, False)
    Call AddRoundBox(sld, 0.8, 4.8, 5.6, 1.8, CLR_WHITE)
    Call AddAccentBar(sld, 0.8, 4.8, 1.8, CLR_RED)
    Call AddTxt(sld, 1.2, 4.95, 5, 0.35, "IMPLICATION", 14, CLR_RED, True)
    Call AddTxt(sld, 1.2, 5.35, 5, 1, _
        "Model must not default to always predicting the majority class. " & _
        "ADASYN oversampling and class-weighting strategies are required.", 13, CLR_DARKTEXT, False)
    Call AddRoundBox(sld, 6.8, 4.8, 5.7, 1.8, CLR_WHITE)
    Call AddAccentBar(sld, 6.8, 4.8, 1.8, CLR_TEAL)
    Call AddTxt(sld, 7.2, 4.95, 5, 0.35, "DATA NOTE", 14, CLR_TEAL, True)
    Call AddTxt(sld, 7.2, 5.35, 5, 1, _
        "After drop_duplicates(), 1,000 duplicate rows removed. " & _
        "100,000 clean records remain. 95/5 ratio preserved.", 13, CLR_DARKTEXT, False)
End Sub

' ════════════════════════════════════════════════════════════
' SLIDE: UNIVARIATE CONTINUOUS
' ════════════════════════════════════════════════════════════
Private Sub Slide_UnivariateContinuous(prs As Presentation)
    Dim sld As Slide, tbl As Table, shp As Shape, r As Integer, c As Integer
    Dim td(1 To 7, 1 To 5) As String
    Set sld = prs.Slides.Add(prs.Slides.Count + 1, ppLayoutBlank)
    Call ContentSlideSetup(sld, "1.2  Univariate Analysis " & ChrW(8212) & " Continuous Features", "Distribution profiles reveal extreme skew in financial variables")
    td(1, 1) = "Feature": td(1, 2) = "Mean": td(1, 3) = "Median": td(1, 4) = "Skewness": td(1, 5) = "Distribution"
    td(2, 1) = "Age": td(2, 2) = "44.51": td(2, 3) = "45.0": td(2, 4) = "-0.002": td(2, 5) = "Symmetric"
    td(3, 1) = "Income": td(3, 2) = "75,907": td(3, 3) = "70,263": td(3, 4) = "10.01": td(3, 5) = "Highly Right-Skewed"
    td(4, 1) = "TotalSpend": td(4, 2) = "5,191": td(4, 3) = "5,029": td(4, 4) = "5.09": td(4, 5) = "Highly Right-Skewed"
    td(5, 1) = "TotalTransactions": td(5, 2) = "99.99": td(5, 3) = "100.0": td(5, 4) = "0.10": td(5, 5) = "Symmetric"
    td(6, 1) = "CreditLimit": td(6, 2) = ChrW(8212): td(6, 3) = "14,964": td(6, 4) = ChrW(8212): td(6, 5) = "Right-Skewed"
    td(7, 1) = "Tenure": td(7, 2) = ChrW(8212): td(7, 3) = "10 yrs": td(7, 4) = ChrW(8212): td(7, 5) = "Near-Uniform"
    Set shp = sld.Shapes.AddTable(7, 5, 0.8 * 72, 1.6 * 72, 11.7 * 72, 3.6 * 72)
    Set tbl = shp.Table
    For r = 1 To 7: For c = 1 To 5
        Call FormatTableCell(tbl, r, c, td(r, c), 13, 7)
        If r = 1 Then tbl.Cell(r, c).Shape.TextFrame2.TextRange.Font.Size = 14
    Next c: Next r
    Call AddRoundBox(sld, 0.8, 5.6, 11.7, 1.2, CLR_WHITE)
    Call AddAccentBar(sld, 0.8, 5.6, 1.2, CLR_GOLD)
    Call AddTxt(sld, 1.2, 5.68, 1.2, 0.35, "INSIGHT", 13, CLR_GOLD, True)
    Call AddTxt(sld, 2.5, 5.68, 9.5, 0.9, _
        "Income and TotalSpend exhibit extreme right skew (>5) with leptokurtic tails. This drives the decision " & _
        "to use median imputation for missing values and winsorization for the Logistic Regression branch.", 13, CLR_DARKTEXT, False)
End Sub

' ════════════════════════════════════════════════════════════
' SLIDE: DATA QUALITY
' ════════════════════════════════════════════════════════════
Private Sub Slide_DataQuality(prs As Presentation)
    Dim sld As Slide, i As Integer, col As Integer, row As Integer, cx As Single, cy As Single
    Dim issues(1 To 6, 1 To 3) As String, issClrs(1 To 6) As Long
    Set sld = prs.Slides.Add(prs.Slides.Count + 1, ppLayoutBlank)
    Call ContentSlideSetup(sld, "1.3  Data Quality Issues Identified", "Systematic audit of anomalies, missing data, and structural challenges")
    issues(1, 1) = "Missing Values": issues(1, 2) = "~5,000 each (~5%)": issues(1, 3) = "Median imputation + _is_missing flag"
    issues(2, 1) = "Negative Income": issues(2, 2) = "23 records": issues(2, 3) = "Flagged as Income_is_negative"
    issues(3, 1) = "Negative CreditLimit": issues(3, 2) = "142 records": issues(3, 3) = "Flagged as CreditLimit_is_negative"
    issues(4, 1) = "Negative TotalSpend": issues(4, 2) = "578 (~0.5%)": issues(4, 3) = "Flagged as TotalSpend_is_negative"
    issues(5, 1) = "Duplicate Records": issues(5, 2) = "1,000 exact": issues(5, 3) = "Dropped via drop_duplicates()"
    issues(6, 1) = "High Cardinality": issues(6, 2) = "Country: 100": issues(6, 3) = "Target Encoding (smoothed)"
    issClrs(1) = CLR_ACCENT: issClrs(2) = CLR_RED: issClrs(3) = CLR_RED
    issClrs(4) = CLR_RED: issClrs(5) = CLR_TEAL: issClrs(6) = CLR_NAVY
    For i = 1 To 6
        col = ((i - 1) Mod 3): row = ((i - 1) \ 3)
        cx = 0.8 + col * 4.05: cy = 1.6 + row * 2.6
        Call AddRoundBox(sld, cx, cy, 3.8, 2.3, CLR_WHITE)
        Call AddBox(sld, cx, cy, 3.8, 0.06, issClrs(i))
        Call AddTxt(sld, cx + 0.3, cy + 0.25, 3.2, 0.35, issues(i, 1), 15, issClrs(i), True)
        Call AddTxt(sld, cx + 0.3, cy + 0.7, 3.2, 0.4, issues(i, 2), 20, CLR_DARK, True, ppAlignLeft, FNT_HEAD)
        Call AddTxt(sld, cx + 0.3, cy + 1.25, 3.2, 0.8, "Action: " & issues(i, 3), 12, CLR_GRAY, False)
    Next i
End Sub

' ════════════════════════════════════════════════════════════
' SLIDE: UNIVARIATE CATEGORICAL
' ════════════════════════════════════════════════════════════
Private Sub Slide_UnivariateCategorical(prs As Presentation)
    Dim sld As Slide, i As Integer, cx As Single, bdg As Shape
    Dim feats(1 To 4, 1 To 4) As String, cW2 As Single, gap2 As Single, sx As Single
    Set sld = prs.Slides.Add(prs.Slides.Count + 1, ppLayoutBlank)
    Call ContentSlideSetup(sld, "1.4  Univariate Analysis " & ChrW(8212) & " Categorical Features", "Chi-square tests reveal NO significant association with attrition")
    feats(1, 1) = "Gender": feats(1, 2) = "Male 5.02%  |  Female 4.91%": feats(1, 3) = "Chi2 p = 0.43": feats(1, 4) = "NOT SIGNIFICANT"
    feats(2, 1) = "MaritalStatus": feats(2, 2) = "Widowed 5.12%  |  Single 4.86%": feats(2, 3) = "Chi2 p = 0.47": feats(2, 4) = "NOT SIGNIFICANT"
    feats(3, 1) = "EducationLevel": feats(3, 2) = "Master 5.12%  |  HS 4.80%": feats(3, 3) = "Chi2 p = 0.33": feats(3, 4) = "NOT SIGNIFICANT"
    feats(4, 1) = "CardType": feats(4, 2) = "Marginal spread across types": feats(4, 3) = "Cramer's V negligible": feats(4, 4) = "NOT SIGNIFICANT"
    cW2 = 2.75: gap2 = 0.2: sx = (SW - (4 * cW2 + 3 * gap2)) / 2
    For i = 1 To 4
        cx = sx + (i - 1) * (cW2 + gap2)
        Call AddRoundBox(sld, cx, 1.7, cW2, 3.2, CLR_WHITE)
        Call AddTxt(sld, cx + 0.25, 1.9, cW2 - 0.5, 0.4, feats(i, 1), 17, CLR_DARK, True)
        Call AddBox(sld, cx + 0.25, 2.35, 1.5, 0.02, CLR_ACCENT)
        Call AddTxt(sld, cx + 0.25, 2.55, cW2 - 0.5, 0.8, feats(i, 2), 12, CLR_DARKTEXT, False)
        Call AddTxt(sld, cx + 0.25, 3.35, cW2 - 0.5, 0.35, feats(i, 3), 12, CLR_GRAY, False)
        Set bdg = AddRoundBox(sld, cx + 0.25, 3.85, cW2 - 0.5, 0.45, CLR_RED)
        Call AddTxt(sld, cx + 0.25, 3.88, cW2 - 0.5, 0.4, feats(i, 4), 12, CLR_WHITE, True, ppAlignCenter)
    Next i
    Call AddRoundBox(sld, 0.8, 5.3, 11.7, 1.4, CLR_WHITE)
    Call AddAccentBar(sld, 0.8, 5.3, 1.4)
    Call AddTxt(sld, 1.2, 5.45, 3, 0.35, "Country (100 values)", 14, CLR_DARK, True)
    Call AddTxt(sld, 1.2, 5.85, 10.5, 0.6, _
        "Target Encoding std dev = 0.0068 (extremely narrow). Zero high-risk countries identified " & _
        "(0 countries exceed 1.5x global churn rate). Geography adds no predictive value.", 13, CLR_DARKTEXT, False)
End Sub

' ════════════════════════════════════════════════════════════
' SLIDE: BIVARIATE ANALYSIS
' ════════════════════════════════════════════════════════════
Private Sub Slide_BivariateAnalysis(prs As Presentation)
    Dim sld As Slide, tbl As Table, shp As Shape, r As Integer, c As Integer
    Dim td(1 To 7, 1 To 5) As String
    Set sld = prs.Slides.Add(prs.Slides.Count + 1, ppLayoutBlank)
    Call ContentSlideSetup(sld, "1.5  Bivariate Analysis " & ChrW(8212) & " Numeric vs. Attrition", "Median comparisons reveal near-zero difference across all features")
    td(1, 1) = "Feature": td(1, 2) = "Median (Stayed)": td(1, 3) = "Median (Attrited)": td(1, 4) = "T-Test p-value": td(1, 5) = "Cohen's d"
    td(2, 1) = "Age": td(2, 2) = "45.0": td(2, 3) = "44.0": td(2, 4) = "p = 0.128": td(2, 5) = "Negligible"
    td(3, 1) = "Income": td(3, 2) = "70,263": td(3, 3) = "70,263": td(3, 4) = "Not Sig.": td(3, 5) = "Negligible"
    td(4, 1) = "CreditLimit": td(4, 2) = "14,964": td(4, 3) = "14,964": td(4, 4) = "Not Sig.": td(4, 5) = "Negligible"
    td(5, 1) = "TotalTransactions": td(5, 2) = "100": td(5, 3) = "100": td(5, 4) = "p = 0.499": td(5, 5) = "d = -0.010"
    td(6, 1) = "TotalSpend": td(6, 2) = "5,029": td(6, 3) = "5,029": td(6, 4) = "p = 0.186": td(6, 5) = "d = +0.020"
    td(7, 1) = "Tenure": td(7, 2) = "10.0 yrs": td(7, 3) = "10.0 yrs": td(7, 4) = "p = 0.104": td(7, 5) = "d = -0.024"
    Set shp = sld.Shapes.AddTable(7, 5, 0.8 * 72, 1.6 * 72, 11.7 * 72, 3.2 * 72)
    Set tbl = shp.Table
    For r = 1 To 7: For c = 1 To 5: Call FormatTableCell(tbl, r, c, td(r, c), 13, 7): Next c: Next r
    Call AddRoundBox(sld, 0.8, 5.2, 11.7, 1.5, CLR_WHITE)
    Call AddAccentBar(sld, 0.8, 5.2, 1.5, CLR_RED)
    Call AddTxt(sld, 1.2, 5.3, 3, 0.35, "KEY FINDING", 14, CLR_RED, True)
    Call AddTxt(sld, 1.2, 5.7, 10.8, 0.8, _
        "All features show statistically identical medians between groups. Cohen's d values are uniformly " & _
        "below 0.1, confirming negligible effect sizes. No raw feature separates the two classes.", 13, CLR_DARKTEXT, False)
End Sub

' ════════════════════════════════════════════════════════════
' SLIDE: CORRELATION MATRIX
' ════════════════════════════════════════════════════════════
Private Sub Slide_CorrelationMatrix(prs As Presentation)
    Dim sld As Slide, i As Integer, col As Integer, row As Integer, cx As Single, cy As Single
    Dim feats As Variant, pearson As Variant, cW3 As Single, gap3 As Single
    Set sld = prs.Slides.Add(prs.Slides.Count + 1, ppLayoutBlank)
    Call ContentSlideSetup(sld, "1.6  Multivariate Analysis " & ChrW(8212) & " Correlation Matrix", "All correlations with AttritionFlag are at noise level")
    feats = Array("Age", "Income", "CreditLimit", "TotalTxns", "TotalSpend", "Tenure", "Avg_Txn/Tenure", "Credit_Util")
    pearson = Array("-0.005", "~0.00", "~0.00", "~0.00", "~0.00", "-0.005", "+0.007", "+0.003")
    cW3 = 2.7: gap3 = 0.15
    For i = 0 To 7
        col = i Mod 4: row = i \ 4
        cx = 0.8 + col * (cW3 + gap3): cy = 1.7 + row * 2.4
        Call AddRoundBox(sld, cx, cy, cW3, 2.1, CLR_WHITE)
        Call AddTxt(sld, cx + 0.2, cy + 0.2, cW3 - 0.4, 0.35, CStr(feats(i)), 14, CLR_DARK, True)
        Call AddTxt(sld, cx + 0.2, cy + 0.65, cW3 - 0.4, 0.5, "Pearson r", 12, CLR_GRAY, False, ppAlignCenter)
        Call AddTxt(sld, cx + 0.2, cy + 1#, cW3 - 0.4, 0.6, CStr(pearson(i)), 28, CLR_RED, True, ppAlignCenter, FNT_HEAD)
        Call AddTxt(sld, cx + 0.2, cy + 1.6, cW3 - 0.4, 0.3, "Noise-level", 12, CLR_GRAY, False, ppAlignCenter)
    Next i
    Call AddRoundBox(sld, 0.8, 6.6, 11.7, 0.65, CLR_DARK)
    Call AddTxt(sld, 1.2, 6.67, 11, 0.45, _
        "Even engineered features yield correlations below |r| = 0.03. Non-linear modeling is required.", 14, CLR_ICE, False, ppAlignCenter)
End Sub

' ════════════════════════════════════════════════════════════
' SLIDE: SECTION 1 SYNTHESIS
' ════════════════════════════════════════════════════════════
Private Sub Slide_Section1_Synthesis(prs As Presentation)
    Dim sld As Slide, findings As String, implications As String
    Set sld = prs.Slides.Add(prs.Slides.Count + 1, ppLayoutBlank)
    Call SetDarkBg(sld): Call AddBox(sld, 0, 0, SW, 0.06, CLR_ACCENT)
    Call AddTxt(sld, 0.8, 0.4, 11, 0.5, "Section 1 " & ChrW(8212) & " Synthesis", 28, CLR_WHITE, True, ppAlignLeft, FNT_HEAD)
    Call AddTxt(sld, 0.8, 0.95, 11, 0.4, "Raw features carry almost no discriminative signal for attrition", 17, CLR_ACCENT, True)
    Call AddRoundBox(sld, 0.8, 1.7, 5.6, 4.8, CLR_CARD)
    Call AddTxt(sld, 1.2, 1.9, 4.8, 0.4, "WHAT THE DATA SHOWS", 15, CLR_ACCENT, True)
    findings = "Continuous features: near-identical distributions" & vbCrLf & vbCrLf & _
               "Categorical features: uniformly balanced" & vbCrLf & vbCrLf & _
               "Correlation matrix: all |r| < 0.01" & vbCrLf & vbCrLf & _
               "Engineered ratios: still weak" & vbCrLf & vbCrLf & _
               "Chi-square: all p > 0.05, Cramer's V negligible"
    Call AddTxt(sld, 1.2, 2.45, 4.8, 3.6, findings, 14, CLR_ICE, False)
    Call AddRoundBox(sld, 6.8, 1.7, 5.7, 4.8, CLR_CARD)
    Call AddTxt(sld, 7.2, 1.9, 5, 0.4, "MODELING IMPLICATIONS", 15, CLR_TEAL, True)
    implications = "Simple rules CANNOT identify churners" & vbCrLf & vbCrLf & _
                   "Feature engineering: necessary but not sufficient" & vbCrLf & vbCrLf & _
                   "Non-linear ensembles (XGBoost, RF) required" & vbCrLf & vbCrLf & _
                   "Class imbalance (95/5) must be handled" & vbCrLf & vbCrLf & _
                   "Target encoding needed for Country (100 cats)"
    Call AddTxt(sld, 7.2, 2.45, 5, 3.6, implications, 14, CLR_ICE, False)
End Sub

' ════════════════════════════════════════════════════════════
' SLIDE: BEHAVIORAL  (minimalist Q&A format)
' ════════════════════════════════════════════════════════════
Private Sub Slide_Behavioral(prs As Presentation)
    Dim sld As Slide
    Dim PHP As String: PHP = ChrW(8369)
    Set sld = prs.Slides.Add(prs.Slides.Count + 1, ppLayoutBlank)
    Call ContentSlideSetup(sld, "2.1  Behavioral & Transactional Indicators", "")

    ' — Question
    Call AddTxt(sld, 0.8, 1.45, 11.7, 0.3, "BUSINESS QUESTION", 10, CLR_ACCENT, True)
    Call AddTxt(sld, 0.8, 1.73, 11.7, 0.4, "Do churners show different transaction or spending patterns before leaving?", 14, CLR_GRAY, False)
    Call AddBox(sld, 0.8, 2.18, 11.7, 0.02, CLR_LTGRAY)

    ' — Headline answer
    Call AddTxt(sld, 0.8, 2.28, 11.7, 0.8, "Churners and stayers are behaviourally indistinguishable.", 30, CLR_DARK, True, ppAlignLeft, FNT_HEAD)
    Call AddBox(sld, 0.8, 3.2, 11.7, 0.02, CLR_LTGRAY)

    ' — Supporting stats (3 columns, no boxes)
    Call AddTxt(sld, 0.8, 3.35, 3.6, 0.28, "TotalTransactions", 10, CLR_GRAY, False)
    Call AddTxt(sld, 0.8, 3.62, 3.6, 0.55, "100 vs 100", 22, CLR_DARK, True, ppAlignLeft, FNT_HEAD)
    Call AddTxt(sld, 0.8, 4.2, 3.6, 0.28, "p = 0.499   |   d = -0.010", 10, CLR_GRAY, False)

    Call AddTxt(sld, 4.9, 3.35, 3.8, 0.28, "TotalSpend  (median)", 10, CLR_GRAY, False)
    Call AddTxt(sld, 4.9, 3.62, 3.8, 0.55, PHP & "5,029 vs " & PHP & "5,029", 22, CLR_DARK, True, ppAlignLeft, FNT_HEAD)
    Call AddTxt(sld, 4.9, 4.2, 3.8, 0.28, "p = 0.186   |   d = +0.020", 10, CLR_GRAY, False)

    Call AddTxt(sld, 9.0, 3.35, 3.5, 0.28, "Credit Utilization", 10, CLR_GRAY, False)
    Call AddTxt(sld, 9.0, 3.62, 3.5, 0.55, "0.336 vs 0.337", 22, CLR_DARK, True, ppAlignLeft, FNT_HEAD)
    Call AddTxt(sld, 9.0, 4.2, 3.5, 0.28, "p = 0.429   |   r = 0.003", 10, CLR_GRAY, False)

    Call AddBox(sld, 0.8, 4.62, 11.7, 0.02, CLR_LTGRAY)

    ' — Additional signal
    Call AddTxt(sld, 0.8, 4.75, 11.7, 0.28, "ADDITIONAL SIGNAL", 10, CLR_ACCENT, True)
    Call AddTxt(sld, 0.8, 5.05, 11.7, 0.55, _
        "Anomaly flags (missing / negative spend, negative credit limit) produce a maximum lift of +0.67 pp" & vbCrLf & _
        "— below the 1 pp actionability threshold. No flag qualifies as a usable predictor.", _
        13, CLR_DARKTEXT, False)
End Sub

' ════════════════════════════════════════════════════════════
' SLIDE: TENURE LIFECYCLE  (minimalist Q&A format)
' ════════════════════════════════════════════════════════════
Private Sub Slide_TenureLifecycle(prs As Presentation)
    Dim sld As Slide
    Set sld = prs.Slides.Add(prs.Slides.Count + 1, ppLayoutBlank)
    Call ContentSlideSetup(sld, "2.2  Tenure & Customer Lifecycle", "")

    ' — Question
    Call AddTxt(sld, 0.8, 1.45, 11.7, 0.3, "BUSINESS QUESTION", 10, CLR_ACCENT, True)
    Call AddTxt(sld, 0.8, 1.73, 11.7, 0.4, "Do customers who are brand new — or have stayed longest — churn at higher rates?", 14, CLR_GRAY, False)
    Call AddBox(sld, 0.8, 2.18, 11.7, 0.02, CLR_LTGRAY)

    ' — Headline answer
    Call AddTxt(sld, 0.8, 2.28, 11.7, 0.8, "There is no danger zone. Long tenure offers no protection.", 30, CLR_DARK, True, ppAlignLeft, FNT_HEAD)
    Call AddBox(sld, 0.8, 3.2, 11.7, 0.02, CLR_LTGRAY)

    ' — Supporting stats
    Call AddTxt(sld, 0.8, 3.35, 3.6, 0.28, "Spearman correlation (tenure vs churn)", 10, CLR_GRAY, False)
    Call AddTxt(sld, 0.8, 3.62, 3.6, 0.55, "r = -0.005", 22, CLR_DARK, True, ppAlignLeft, FNT_HEAD)
    Call AddTxt(sld, 0.8, 4.2, 3.6, 0.28, "Effectively zero — no linear pattern exists", 10, CLR_GRAY, False)

    Call AddTxt(sld, 4.9, 3.35, 3.8, 0.28, "Highest single-year churn rate", 10, CLR_GRAY, False)
    Call AddTxt(sld, 4.9, 3.62, 3.8, 0.55, "5.71%  (Year 1)", 22, CLR_DARK, True, ppAlignLeft, FNT_HEAD)
    Call AddTxt(sld, 4.9, 4.2, 3.8, 0.28, "+0.75 pp above baseline — not actionable", 10, CLR_GRAY, False)

    Call AddTxt(sld, 9.0, 3.35, 3.5, 0.28, "Global baseline", 10, CLR_GRAY, False)
    Call AddTxt(sld, 9.0, 3.62, 3.5, 0.55, "4.96%", 22, CLR_DARK, True, ppAlignLeft, FNT_HEAD)
    Call AddTxt(sld, 9.0, 4.2, 3.5, 0.28, "50% of all churns occur by Year 10", 10, CLR_GRAY, False)

    Call AddBox(sld, 0.8, 4.62, 11.7, 0.02, CLR_LTGRAY)

    Call AddTxt(sld, 0.8, 4.75, 11.7, 0.28, "IMPLICATION", 10, CLR_ACCENT, True)
    Call AddTxt(sld, 0.8, 5.05, 11.7, 0.55, _
        "No early-warning signal exists in tenure data. The median cohort sits at the exact global baseline." & vbCrLf & _
        "Tenure alone cannot serve as a trigger for retention outreach.", _
        13, CLR_DARKTEXT, False)
End Sub

' ════════════════════════════════════════════════════════════
' SLIDE: FINANCIAL DEMOGRAPHIC  (minimalist Q&A format)
' ════════════════════════════════════════════════════════════
Private Sub Slide_FinancialDemographic(prs As Presentation)
    Dim sld As Slide
    Set sld = prs.Slides.Add(prs.Slides.Count + 1, ppLayoutBlank)
    Call ContentSlideSetup(sld, "2.3  Financial & Demographic Profiles", "")

    ' — Question
    Call AddTxt(sld, 0.8, 1.45, 11.7, 0.3, "BUSINESS QUESTION", 10, CLR_ACCENT, True)
    Call AddTxt(sld, 0.8, 1.73, 11.7, 0.4, "Are wealthier clients, older clients, or premium cardholders less likely to churn?", 14, CLR_GRAY, False)
    Call AddBox(sld, 0.8, 2.18, 11.7, 0.02, CLR_LTGRAY)

    ' — Headline answer
    Call AddTxt(sld, 0.8, 2.28, 11.7, 0.8, "Income, age, and card type produce a perfectly flat churn rate.", 30, CLR_DARK, True, ppAlignLeft, FNT_HEAD)
    Call AddBox(sld, 0.8, 3.2, 11.7, 0.02, CLR_LTGRAY)

    ' — Supporting stats
    Call AddTxt(sld, 0.8, 3.35, 3.6, 0.28, "Income quintiles  (Q1 to Q5)", 10, CLR_GRAY, False)
    Call AddTxt(sld, 0.8, 3.62, 3.6, 0.55, "All ~4.96%", 22, CLR_DARK, True, ppAlignLeft, FNT_HEAD)
    Call AddTxt(sld, 0.8, 4.2, 3.6, 0.28, "Zero gradient from lowest to highest income", 10, CLR_GRAY, False)

    Call AddTxt(sld, 4.9, 3.35, 3.8, 0.28, "Age bands  (<30 to 60+)", 10, CLR_GRAY, False)
    Call AddTxt(sld, 4.9, 3.62, 3.8, 0.55, "All ~4.96%", 22, CLR_DARK, True, ppAlignLeft, FNT_HEAD)
    Call AddTxt(sld, 4.9, 4.2, 3.8, 0.28, "No lifecycle dip or peak in any age group", 10, CLR_GRAY, False)

    Call AddTxt(sld, 9.0, 3.35, 3.5, 0.28, "Card tier  (Blue / Silver / Gold / Platinum)", 10, CLR_GRAY, False)
    Call AddTxt(sld, 9.0, 3.62, 3.5, 0.55, "All ~4.96%", 22, CLR_DARK, True, ppAlignLeft, FNT_HEAD)
    Call AddTxt(sld, 9.0, 4.2, 3.5, 0.28, "Premium cardholders churn at the same rate", 10, CLR_GRAY, False)

    Call AddBox(sld, 0.8, 4.62, 11.7, 0.02, CLR_LTGRAY)

    Call AddTxt(sld, 0.8, 4.75, 11.7, 0.28, "IMPLICATION", 10, CLR_ACCENT, True)
    Call AddTxt(sld, 0.8, 5.05, 11.7, 0.55, _
        "Neither wealth nor age nor card tier predicts attrition. Standard demographic segmentation has no retention value here." & vbCrLf & _
        "All quintiles and bands cluster at the 4.96% global baseline.", _
        13, CLR_DARKTEXT, False)
End Sub

' ════════════════════════════════════════════════════════════
' SLIDE: GEOGRAPHIC SOCIAL  (minimalist Q&A format)
' ════════════════════════════════════════════════════════════
Private Sub Slide_GeographicSocial(prs As Presentation)
    Dim sld As Slide
    Set sld = prs.Slides.Add(prs.Slides.Count + 1, ppLayoutBlank)
    Call ContentSlideSetup(sld, "2.4  Social & Geographic Factors", "")

    ' — Question
    Call AddTxt(sld, 0.8, 1.45, 11.7, 0.3, "BUSINESS QUESTION", 10, CLR_ACCENT, True)
    Call AddTxt(sld, 0.8, 1.73, 11.7, 0.4, "Do marital status, education, gender, or country of origin predict who will churn?", 14, CLR_GRAY, False)
    Call AddBox(sld, 0.8, 2.18, 11.7, 0.02, CLR_LTGRAY)

    ' — Headline answer
    Call AddTxt(sld, 0.8, 2.28, 11.7, 0.8, "Zero high-risk countries. Zero significant demographics. Geography is noise.", 30, CLR_DARK, True, ppAlignLeft, FNT_HEAD)
    Call AddBox(sld, 0.8, 3.2, 11.7, 0.02, CLR_LTGRAY)

    ' — 4 findings in 2 rows of 2 (no boxes)
    Call AddTxt(sld, 0.8, 3.35, 5.8, 0.28, "Gender", 10, CLR_GRAY, False)
    Call AddTxt(sld, 0.8, 3.62, 5.8, 0.42, "Chi" & ChrW(178) & " = 0.617   p = 0.432   V = 0.0025", 14, CLR_DARK, True, ppAlignLeft, FNT_HEAD)
    Call AddTxt(sld, 0.8, 4.07, 5.8, 0.28, "Not significant", 10, CLR_GRAY, False)

    Call AddTxt(sld, 7.1, 3.35, 5.5, 0.28, "MaritalStatus", 10, CLR_GRAY, False)
    Call AddTxt(sld, 7.1, 3.62, 5.5, 0.42, "Chi" & ChrW(178) & " = 2.543   p = 0.468   V = 0.0050", 14, CLR_DARK, True, ppAlignLeft, FNT_HEAD)
    Call AddTxt(sld, 7.1, 4.07, 5.5, 0.28, "Not significant", 10, CLR_GRAY, False)

    Call AddBox(sld, 0.8, 4.48, 11.7, 0.02, CLR_LTGRAY)

    Call AddTxt(sld, 0.8, 4.6, 5.8, 0.28, "EducationLevel", 10, CLR_GRAY, False)
    Call AddTxt(sld, 0.8, 4.87, 5.8, 0.42, "Chi" & ChrW(178) & " = 3.394   p = 0.335   V = 0.0058", 14, CLR_DARK, True, ppAlignLeft, FNT_HEAD)
    Call AddTxt(sld, 0.8, 5.32, 5.8, 0.28, "Not significant", 10, CLR_GRAY, False)

    Call AddTxt(sld, 7.1, 4.6, 5.5, 0.28, "Country  (100 categories, Target-Encoded)", 10, CLR_GRAY, False)
    Call AddTxt(sld, 7.1, 4.87, 5.5, 0.42, "TE std dev = 0.0068   |   0 high-risk nations", 14, CLR_DARK, True, ppAlignLeft, FNT_HEAD)
    Call AddTxt(sld, 7.1, 5.32, 5.5, 0.28, "Encoding range collapsed to noise level", 10, CLR_GRAY, False)

    Call AddBox(sld, 0.8, 5.72, 11.7, 0.02, CLR_LTGRAY)
    Call AddTxt(sld, 0.8, 5.85, 11.7, 0.4, _
        "All social and geographic attributes fail to separate churners from non-churners. These features were retained for model completeness, not signal.", _
        12, CLR_GRAY, False)
End Sub

' ════════════════════════════════════════════════════════════
' SLIDE: SECTION 2 SYNTHESIS  (minimalist format)
' ════════════════════════════════════════════════════════════
Private Sub Slide_Section2_Synthesis(prs As Presentation)
    Dim sld As Slide, i As Integer
    Dim findings(1 To 4, 1 To 2) As String
    Set sld = prs.Slides.Add(prs.Slides.Count + 1, ppLayoutBlank)
    Call SetDarkBg(sld): Call AddBox(sld, 0, 0, SW, 0.06, CLR_ACCENT)
    Call AddTxt(sld, 0.8, 0.4, 11, 0.5, "Section 2 " & ChrW(8212) & " EDA & Feature Discovery", 28, CLR_WHITE, True, ppAlignLeft, FNT_HEAD)

    ' — Verdict
    Call AddTxt(sld, 0.8, 1.2, 11.7, 0.28, "THE VERDICT", 10, CLR_ACCENT, True)
    Call AddBox(sld, 0.8, 1.55, 11.7, 0.02, CLR_CARD)
    Call AddTxt(sld, 0.8, 1.68, 11.7, 0.7, "No feature meaningfully separates attriters from non-attriters.", 26, CLR_WHITE, True, ppAlignLeft, FNT_HEAD)
    Call AddBox(sld, 0.8, 2.52, 11.7, 0.02, CLR_CARD)

    ' — Per-chapter findings
    findings(1, 1) = "Behavioral (2.1)"
    findings(1, 2) = "Transaction volume and spending are statistically identical across groups  (p > 0.05 for all metrics)"
    findings(2, 1) = "Lifecycle (2.2)"
    findings(2, 2) = "No tenure danger zone — highest single-year rate is +0.75 pp above baseline  (Spearman r = -0.005)"
    findings(3, 1) = "Financial (2.3)"
    findings(3, 2) = "All income quintiles, age bands, and card tiers flat at ~4.96% — no wealth or age gradient exists"
    findings(4, 1) = "Geographic (2.4)"
    findings(4, 2) = "0 high-risk countries — TE std dev = 0.0068 (noise level).  All Cram" & ChrW(233) & "r's V < 0.006"

    For i = 1 To 4
        Dim cy As Single: cy = 2.75 + (i - 1) * 1.0
        Call AddTxt(sld, 0.8, cy, 3.2, 0.35, findings(i, 1), 13, CLR_ACCENT, True)
        Call AddTxt(sld, 4.1, cy, 8.5, 0.35, findings(i, 2), 13, CLR_ICE, False)
        If i < 4 Then Call AddBox(sld, 0.8, cy + 0.55, 11.7, 0.01, CLR_CARD)
    Next i

    Call AddBox(sld, 0.8, 6.82, 11.7, 0.02, CLR_ACCENT)
    Call AddTxt(sld, 0.8, 6.95, 11.7, 0.35, _
        "Modeling is not optional here — simple rules and demographics cannot identify who will churn.", _
        12, CLR_ACCENT, False, ppAlignCenter)
End Sub

' ════════════════════════════════════════════════════════════
' SLIDE: MODEL SETUP (Section 3)
' ════════════════════════════════════════════════════════════
Private Sub Slide_ModelSetup(prs As Presentation)
    Dim sld As Slide, dsInfo As String, feInfo As String
    Set sld = prs.Slides.Add(prs.Slides.Count + 1, ppLayoutBlank)
    Call ContentSlideSetup(sld, "3.1  Setup & Data Configuration", "Full 63-column dataset with feature engineering applied")
    Call AddRoundBox(sld, 0.8, 1.7, 5.8, 4.5, CLR_WHITE)
    Call AddAccentBar(sld, 0.8, 1.7, 4.5)
    Call AddTxt(sld, 1.2, 1.9, 5, 0.35, "DATASET", 15, CLR_ACCENT, True)
    dsInfo = "Full 63-column dataset (incl. Feature_0 to Feature_49)" & vbCrLf & vbCrLf & _
             "101K raw " & ChrW(8594) & " 100K post-dedup" & vbCrLf & vbCrLf & _
             "Churn rate: 4.96% (4,960 attrited)" & vbCrLf & vbCrLf & _
             "Stratified 70/30 train-test split" & vbCrLf & vbCrLf & _
             "Decision Tree: class_weight='balanced'" & vbCrLf & "XGBoost: scale_pos_weight = 19.14"
    Call AddTxt(sld, 1.2, 2.4, 5, 3.5, dsInfo, 14, CLR_DARKTEXT, False)
    Call AddRoundBox(sld, 7, 1.7, 5.5, 4.5, CLR_WHITE)
    Call AddAccentBar(sld, 7, 1.7, 4.5, CLR_TEAL)
    Call AddTxt(sld, 7.4, 1.9, 4.8, 0.35, "FEATURE ENGINEERING", 15, CLR_TEAL, True)
    feInfo = "SpendToLimit = TotalSpend / CreditLimit" & vbCrLf & vbCrLf & _
             "TxPerTenure = TotalTransactions / Tenure" & vbCrLf & vbCrLf & _
             "SpendPerTx = TotalSpend / TotalTransactions" & vbCrLf & vbCrLf & _
             "log1p transforms: Income, CreditLimit, TotalSpend, TotalTransactions" & vbCrLf & vbCrLf & _
             "Interaction terms: Age" & ChrW(215) & "Tenure, Income" & ChrW(215) & "CreditLimit"
    Call AddTxt(sld, 7.4, 2.4, 4.8, 3.5, feInfo, 14, CLR_DARKTEXT, False)
End Sub

' ════════════════════════════════════════════════════════════
' SLIDE: DECISION TREE
' ════════════════════════════════════════════════════════════
Private Sub Slide_DecisionTree(prs As Presentation)
    Dim sld As Slide, i As Integer, cx As Single
    Dim cfgs(1 To 3, 1 To 3) As String, aucClrs(1 To 3) As Long
    Set sld = prs.Slides.Add(prs.Slides.Count + 1, ppLayoutBlank)
    Call ContentSlideSetup(sld, "3.2  Shallow Decision Tree", "Three depth configurations tested  |  Purpose: find simple split rules")
    cfgs(1, 1) = "Depth 2": cfgs(1, 2) = "0.489": cfgs(1, 3) = "Below random"
    cfgs(2, 1) = "Depth 3": cfgs(2, 2) = "0.496": cfgs(2, 3) = "Effectively random"
    cfgs(3, 1) = "Depth 4": cfgs(3, 2) = "0.503": cfgs(3, 3) = "Best (barely)"
    aucClrs(1) = CLR_RED: aucClrs(2) = CLR_RED: aucClrs(3) = CLR_GOLD
    For i = 1 To 3
        cx = 0.8 + (i - 1) * 4.05
        Call AddRoundBox(sld, cx, 1.7, 3.8, 1.8, CLR_WHITE)
        Call AddTxt(sld, cx + 0.3, 1.85, 3.2, 0.35, cfgs(i, 1), 15, CLR_DARK, True, ppAlignCenter)
        Call AddTxt(sld, cx + 0.3, 2.25, 3.2, 0.6, "AUC " & cfgs(i, 2), 28, aucClrs(i), True, ppAlignCenter, FNT_HEAD)
        Call AddTxt(sld, cx + 0.3, 2.9, 3.2, 0.3, cfgs(i, 3), 13, CLR_GRAY, False, ppAlignCenter)
    Next i
    Call AddTxt(sld, 0.8, 3.8, 8, 0.35, "BEST TREE (DEPTH 4) " & ChrW(8212) & " FEATURE IMPORTANCE", 15, CLR_DARK, True)
    Call AddRoundBox(sld, 0.8, 4.3, 11.7, 0.5, CLR_OFFWHITE)
    Call AddBox(sld, 0.8, 4.3, 9.64, 0.5, CLR_ACCENT)
    Call AddTxt(sld, 1.1, 4.33, 5, 0.4, "log1p_Income  82.4%", 14, CLR_WHITE, True)
    Call AddRoundBox(sld, 0.8, 4.95, 11.7, 0.5, CLR_OFFWHITE)
    Call AddBox(sld, 0.8, 4.95, 2.06, 0.5, CLR_TEAL)
    Call AddTxt(sld, 1.1, 4.98, 5, 0.4, "log1p_TotalSpend  17.6%", 14, CLR_WHITE, True)
    Call AddRoundBox(sld, 0.8, 5.6, 11.7, 0.5, CLR_OFFWHITE)
    Call AddTxt(sld, 1.1, 5.63, 5, 0.4, "All other features  0.0%", 14, CLR_GRAY, True)
    Call AddRoundBox(sld, 0.8, 6.4, 11.7, 0.8, CLR_WHITE)
    Call AddAccentBar(sld, 0.8, 6.4, 0.8, CLR_RED)
    Call AddTxt(sld, 1.2, 6.5, 11, 0.5, _
        "Best leaf: 327 customers at 1.79x lift " & ChrW(8212) & " too small for a bank-scale retention campaign.", 14, CLR_DARKTEXT, False)
End Sub

' ════════════════════════════════════════════════════════════
' SLIDE: XGBOOST
' ════════════════════════════════════════════════════════════
Private Sub Slide_XGBoost(prs As Presentation)
    Dim sld As Slide, shp As Shape, tbl As Table, r As Integer, c As Integer
    Dim td(1 To 5, 1 To 4) As String
    Set sld = prs.Slides.Add(prs.Slides.Count + 1, ppLayoutBlank)
    Call ContentSlideSetup(sld, "3.3  XGBoost " & ChrW(8212) & " Full Feature Set", "4,000 rounds with early stopping  |  scale_pos_weight = 19.14")
    Call AddRoundBox(sld, 0.8, 1.7, 4, 3, CLR_WHITE)
    Call AddTxt(sld, 1#, 1.85, 3.6, 0.3, "TEST ROC-AUC", 14, CLR_GRAY, False, ppAlignCenter)
    Call AddTxt(sld, 1#, 2.2, 3.6, 1, "0.499", 56, CLR_RED, True, ppAlignCenter, FNT_HEAD)
    Call AddTxt(sld, 1#, 3.3, 3.6, 0.3, "Effectively a coin flip", 14, CLR_RED, False, ppAlignCenter)
    td(1, 1) = "Decile": td(1, 2) = "Attrition Rate": td(1, 3) = "Avg Pred Prob": td(1, 4) = "Lift"
    td(2, 1) = "9 (Top)": td(2, 2) = "4.79%": td(2, 3) = "0.549": td(2, 4) = "0.96x"
    td(3, 1) = "8": td(3, 2) = "5.05%": td(3, 3) = "0.517": td(3, 4) = "1.02x"
    td(4, 1) = "5": td(4, 2) = "5.18%": td(4, 3) = "0.472": td(4, 4) = "1.04x"
    td(5, 1) = "1 (Low)": td(5, 2) = "4.98%": td(5, 3) = "0.354": td(5, 4) = "1.00x"
    Set shp = sld.Shapes.AddTable(5, 4, 5.2 * 72, 1.7 * 72, 7.3 * 72, 2.7 * 72)
    Set tbl = shp.Table
    For r = 1 To 5: For c = 1 To 4: Call FormatTableCell(tbl, r, c, td(r, c), 13, 5): Next c: Next r
    Call AddRoundBox(sld, 0.8, 5.2, 11.7, 1.5, CLR_WHITE)
    Call AddAccentBar(sld, 0.8, 5.2, 1.5, CLR_RED)
    Call AddTxt(sld, 1.2, 5.35, 11, 0.4, "VERDICT", 15, CLR_RED, True)
    Call AddTxt(sld, 1.2, 5.8, 11, 0.7, _
        "Decile lifts oscillate between 0.94x and 1.04x " & ChrW(8212) & " no decile meaningfully concentrates attrition. " & _
        "The model cannot rank customers by attrition risk.", 14, CLR_DARKTEXT, False)
End Sub

' ════════════════════════════════════════════════════════════
' SLIDE: XGBOOST FEATURE IMPORTANCE
' ════════════════════════════════════════════════════════════
Private Sub Slide_XGBoostFeatureImportance(prs As Presentation)
    Dim sld As Slide, i As Integer, cy As Single, barW As Single, barClr As Long, gainVal As Single
    Dim feats(1 To 7, 1 To 3) As String
    Set sld = prs.Slides.Add(prs.Slides.Count + 1, ppLayoutBlank)
    Call ContentSlideSetup(sld, "3.4  XGBoost Feature Importance (Gain)", "Model latches onto high-cardinality noise and statistically insignificant features")
    feats(1, 1) = "Country_19": feats(1, 2) = "153.9": feats(1, 3) = "Geographic dummy (noise)"
    feats(2, 1) = "Country_22": feats(2, 2) = "117.6": feats(2, 3) = "Geographic dummy (noise)"
    feats(3, 1) = "Country_15": feats(3, 2) = "116.6": feats(3, 3) = "Geographic dummy (noise)"
    feats(4, 1) = "Edu_HighSchool": feats(4, 2) = "114.8": feats(4, 3) = "Chi2 p > 0.33"
    feats(5, 1) = "Card_Platinum": feats(5, 2) = "109.6": feats(5, 3) = "Marginal at best"
    feats(6, 1) = "Feature_28": feats(6, 2) = "103.6": feats(6, 3) = "Anonymous feature"
    feats(7, 1) = "Tenure": feats(7, 2) = "97.8": feats(7, 3) = "Spearman r = -0.005"
    For i = 1 To 7
        cy = 1.55 + (i - 1) * 0.75
        Call AddTxt(sld, 0.8, cy + 0.02, 2.2, 0.35, feats(i, 1), 13, CLR_DARK, True)
        gainVal = CSng(feats(i, 2)): barW = (gainVal / 153.9) * 7
        If i <= 3 Then barClr = CLR_RED Else barClr = CLR_ACCENT
        Call AddRoundBox(sld, 3, cy, barW, 0.4, barClr)
        Call AddTxt(sld, 3 + barW + 0.15, cy + 0.02, 1, 0.35, feats(i, 2), 13, CLR_DARK, True)
        Call AddTxt(sld, 10.5, cy + 0.02, 2.5, 0.35, feats(i, 3), 12, CLR_GRAY, False)
    Next i
    Call AddRoundBox(sld, 0.8, 6.3, 11.7, 0.9, CLR_DARK)
    Call AddTxt(sld, 1.2, 6.4, 11, 0.65, _
        "Top features are country dummies (noise) and insignificant categoricals. Consistent with AUC = 0.499.", 14, CLR_ICE, False, ppAlignCenter)
End Sub

' ════════════════════════════════════════════════════════════
' SLIDE: SECTION 3 MODEL SUMMARY
' ════════════════════════════════════════════════════════════
Private Sub Slide_ModelSummary(prs As Presentation)
    Dim sld As Slide, i As Integer, cx As Single, models(1 To 4, 1 To 4) As String
    Set sld = prs.Slides.Add(prs.Slides.Count + 1, ppLayoutBlank)
    Call SetDarkBg(sld): Call AddBox(sld, 0, 0, SW, 0.06, CLR_ACCENT)
    Call AddTxt(sld, 0.8, 0.4, 11, 0.5, "Section 3 " & ChrW(8212) & " Model Validation Summary", 28, CLR_WHITE, True, ppAlignLeft, FNT_HEAD)
    models(1, 1) = "DT (d=2)": models(1, 2) = "0.489": models(1, 3) = "None": models(1, 4) = "Below random"
    models(2, 1) = "DT (d=3)": models(2, 2) = "0.496": models(2, 3) = "Marginal": models(2, 4) = "Effectively random"
    models(3, 1) = "DT (d=4)": models(3, 2) = "0.503": models(3, 3) = "1.79x (n=327)": models(3, 4) = "Leaf too small"
    models(4, 1) = "XGBoost (4K)": models(4, 2) = "0.499": models(4, 3) = "1.04x": models(4, 4) = "No usable signal"
    For i = 1 To 4
        cx = 0.8 + (i - 1) * 3.1
        Call AddRoundBox(sld, cx, 1.3, 2.85, 3.2, CLR_CARD)
        Call AddTxt(sld, cx + 0.2, 1.5, 2.45, 0.5, models(i, 1), 14, CLR_ICE, True, ppAlignCenter)
        Call AddTxt(sld, cx + 0.2, 2.1, 2.45, 0.7, models(i, 2), 36, CLR_RED, True, ppAlignCenter, FNT_HEAD)
        Call AddTxt(sld, cx + 0.2, 2.8, 2.45, 0.3, "AUC", 13, CLR_GRAY, False, ppAlignCenter)
        Call AddBox(sld, cx + 0.4, 3.2, 2.05, 0.01, CLR_GRAY)
        Call AddTxt(sld, cx + 0.2, 3.35, 2.45, 0.3, "Lift: " & models(i, 3), 13, CLR_ICE, False, ppAlignCenter)
        Call AddTxt(sld, cx + 0.2, 3.7, 2.45, 0.4, models(i, 4), 12, CLR_GRAY, False, ppAlignCenter)
    Next i
End Sub

' ════════════════════════════════════════════════════════════
' SLIDE: PREPROCESSING PIPELINE (Section 4)
' ════════════════════════════════════════════════════════════
Private Sub Slide_PreprocessPipeline(prs As Presentation)
    Dim sld As Slide, lrInfo As String, trInfo As String
    Set sld = prs.Slides.Add(prs.Slides.Count + 1, ppLayoutBlank)
    Call ContentSlideSetup(sld, "4.1" & ChrW(8211) & "4.8  Pipeline Architecture", _
        "Clean once, branch twice  |  80/20 stratified split  |  ADASYN on train only")
    Call AddRoundBox(sld, 4.5, 1.5, 4.3, 1.4, CLR_WHITE)
    Call AddBox(sld, 4.5, 1.5, 4.3, 0.07, CLR_TEAL)
    Call AddTxt(sld, 4.8, 1.7, 3.7, 0.35, "df_base  (100K rows)", 16, CLR_TEAL, True, ppAlignCenter)
    Call AddTxt(sld, 4.8, 2.1, 3.7, 0.6, "Dedup " & ChrW(8594) & " Median impute " & ChrW(8594) & " Negative flags", 13, CLR_DARKTEXT, False, ppAlignCenter)
    Call AddBox(sld, 4.2, 3.05, 2.2, 0.03, CLR_GRAY)
    Call AddBox(sld, 7, 3.05, 2.2, 0.03, CLR_GRAY)
    Call AddRoundBox(sld, 0.8, 3.4, 5.5, 3.4, CLR_WHITE)
    Call AddBox(sld, 0.8, 3.4, 0.08, 3.4, CLR_ACCENT)
    Call AddTxt(sld, 1.2, 3.55, 4.8, 0.4, "df_lr  (LR + EBM)", 16, CLR_ACCENT, True)
    lrInfo = "Winsorized (IQR 1.5x clip)" & vbCrLf & vbCrLf & _
             "2 engineered features kept:" & vbCrLf & _
             "  Credit_Utilization, Avg_Txn/Tenure" & vbCrLf & vbCrLf & _
             "OHE drop_first=True (26 cols)" & vbCrLf & _
             "Target Encoding for Country" & vbCrLf & _
             "StandardScaler applied"
    Call AddTxt(sld, 1.2, 4.05, 4.8, 2.6, lrInfo, 13, CLR_DARKTEXT, False)
    Call AddRoundBox(sld, 7, 3.4, 5.5, 3.4, CLR_WHITE)
    Call AddBox(sld, 7, 3.4, 0.08, 3.4, CLR_TEAL)
    Call AddTxt(sld, 7.4, 3.55, 4.8, 0.4, "df_tree  (XGBoost)", 16, CLR_TEAL, True)
    trInfo = "Raw outliers retained" & vbCrLf & vbCrLf & _
             "All 6 engineered features kept" & vbCrLf & vbCrLf & _
             "OHE full (no drop_first) " & ChrW(8594) & " 34 cols" & vbCrLf & _
             "Target Encoding for Country" & vbCrLf & _
             "No scaling (rank-based splits)" & vbCrLf & _
             "ADASYN: 80K " & ChrW(8594) & " 151,704 rows (50/50)"
    Call AddTxt(sld, 7.4, 4.05, 4.8, 2.6, trInfo, 13, CLR_DARKTEXT, False)
End Sub

' ════════════════════════════════════════════════════════════
' SLIDE: PREPROCESS SUMMARY
' ════════════════════════════════════════════════════════════
Private Sub Slide_PreprocessSummary(prs As Presentation)
    Dim sld As Slide, shp As Shape, tbl As Table, r As Integer, c As Integer
    Dim td(1 To 4, 1 To 6) As String
    Set sld = prs.Slides.Add(prs.Slides.Count + 1, ppLayoutBlank)
    Call ContentSlideSetup(sld, "4.11  Pipeline Summary", _
        "All tables share the same train_idx / test_idx  |  Apples-to-apples comparison")
    td(1, 1) = "Table": td(1, 2) = "Models": td(1, 3) = "Base": td(1, 4) = "Engineered": td(1, 5) = "Encoding": td(1, 6) = "Scaling"
    td(2, 1) = "df_base": td(2, 2) = "Reference": td(2, 3) = "Clean only": td(2, 4) = "None": td(2, 5) = "None": td(2, 6) = "None"
    td(3, 1) = "df_lr": td(3, 2) = "LR, EBM": td(3, 3) = "Winsorized": td(3, 4) = "2 features": td(3, 5) = "OHE + TE": td(3, 6) = "StandardScaler"
    td(4, 1) = "df_tree": td(4, 2) = "XGBoost": td(4, 3) = "Raw outliers": td(4, 4) = "All 6": td(4, 5) = "OHE full + TE": td(4, 6) = "None"
    Set shp = sld.Shapes.AddTable(4, 6, 0.8 * 72, 1.6 * 72, 11.7 * 72, 2.4 * 72)
    Set tbl = shp.Table
    For r = 1 To 4: For c = 1 To 6: Call FormatTableCell(tbl, r, c, td(r, c), 14, 4): Next c: Next r
    Call AddTxt(sld, 0.8, 4.4, 5, 0.35, "SPLIT VERIFICATION", 15, CLR_DARK, True)
    Call AddRoundBox(sld, 0.8, 4.9, 11.7, 1.6, CLR_WHITE)
    Call AddStat(sld, 1.5, 5, 2.2, "80,000", "Train set", CLR_DARK)
    Call AddStat(sld, 4, 5, 2.2, "20,000", "Test set", CLR_DARK)
    Call AddStat(sld, 6.5, 5, 2.5, "4.96%", "Churn rate (both sets)", CLR_RED)
    Call AddStat(sld, 9.3, 5, 2.8, "151,704", "Post-ADASYN train (tree)", CLR_TEAL)
End Sub

' ════════════════════════════════════════════════════════════
' SLIDE: MODEL STRATEGY (Section 5)
' ════════════════════════════════════════════════════════════
Private Sub Slide_ModelStrategy(prs As Presentation)
    Dim sld As Slide, lrDesc As String, ebmDesc As String, xgbDesc As String
    Set sld = prs.Slides.Add(prs.Slides.Count + 1, ppLayoutBlank)
    Call ContentSlideSetup(sld, "5.0  Model Selection Strategy", _
        "Three models spanning the interpretability" & ChrW(8211) & "flexibility spectrum")
    Call AddRoundBox(sld, 0.8, 1.7, 3.7, 4.5, CLR_WHITE)
    Call AddBox(sld, 0.8, 1.7, 3.7, 0.07, CLR_ACCENT)
    Call AddTxt(sld, 1.1, 1.95, 3.2, 0.4, "Logistic Regression", 16, CLR_ACCENT, True)
    Call AddTxt(sld, 1.1, 2.35, 3.2, 0.3, "Interpretability Anchor", 13, CLR_GRAY, False)
    lrDesc = "SMOTE inside imblearn pipeline" & vbCrLf & vbCrLf & _
             "GridSearchCV (C, penalty, solver)" & vbCrLf & _
             "Optimized for Recall" & vbCrLf & vbCrLf & _
             "6 assumptions checked:" & vbCrLf & _
             "Binary, Independent, VIF," & vbCrLf & _
             "Box-Tidwell, Outliers, EPV"
    Call AddTxt(sld, 1.1, 2.8, 3.2, 3.2, lrDesc, 13, CLR_DARKTEXT, False)
    Call AddRoundBox(sld, 4.8, 1.7, 3.7, 4.5, CLR_WHITE)
    Call AddBox(sld, 4.8, 1.7, 3.7, 0.07, CLR_TEAL)
    Call AddTxt(sld, 5.1, 1.95, 3.2, 0.4, "EBM", 16, CLR_TEAL, True)
    Call AddTxt(sld, 5.1, 2.35, 3.2, 0.3, "Glass-Box Bridge", 13, CLR_GRAY, False)
    ebmDesc = "GAM + pairwise interactions" & vbCrLf & _
              "(Microsoft Research)" & vbCrLf & vbCrLf & _
              "RandomizedSearchCV (100 iter)" & vbCrLf & _
              "Handles Box-Tidwell violations" & vbCrLf & vbCrLf & _
              "Per-feature shape functions" & vbCrLf & _
              "Best CV AUC: 0.5414"
    Call AddTxt(sld, 5.1, 2.8, 3.2, 3.2, ebmDesc, 13, CLR_DARKTEXT, False)
    Call AddRoundBox(sld, 8.8, 1.7, 3.7, 4.5, CLR_WHITE)
    Call AddBox(sld, 8.8, 1.7, 3.7, 0.07, CLR_RED)
    Call AddTxt(sld, 9.1, 1.95, 3.2, 0.4, "XGBoost", 16, CLR_RED, True)
    Call AddTxt(sld, 9.1, 2.35, 3.2, 0.3, "Performance Ceiling", 13, CLR_GRAY, False)
    xgbDesc = "RandomizedSearchCV (60 iter)" & vbCrLf & _
              "scale_pos_weight = 19.16" & vbCrLf & vbCrLf & _
              "All 6 engineered features" & vbCrLf & _
              "ADASYN-balanced training" & vbCrLf & vbCrLf & _
              "max_depth=3, lr=0.01" & vbCrLf & _
              "Best CV Recall: 0.1932"
    Call AddTxt(sld, 9.1, 2.8, 3.2, 3.2, xgbDesc, 13, CLR_DARKTEXT, False)
    Call AddRoundBox(sld, 0.8, 6.5, 11.7, 0.65, CLR_DARK)
    Call AddTxt(sld, 1.2, 6.57, 11, 0.45, _
        "Primary: PR-AUC (imbalance-aware)  |  Secondary: F2-Score (Recall-weighted)  |  Business: missed churner > false alarm", _
        13, CLR_ICE, False, ppAlignCenter)
End Sub

' ════════════════════════════════════════════════════════════
' SLIDE: MODEL RESULTS
' ════════════════════════════════════════════════════════════
Private Sub Slide_ModelResults(prs As Presentation)
    Dim sld As Slide, i As Integer, cx As Single
    Dim mdls(1 To 3, 1 To 5) As String, mClrs(1 To 3) As Long
    Set sld = prs.Slides.Add(prs.Slides.Count + 1, ppLayoutBlank)
    Call ContentSlideSetup(sld, "5.1" & ChrW(8211) & "5.3  Individual Model Results", _
        "All models evaluated at F2-optimal thresholds on the same held-out test set")
    mdls(1, 1) = "Logistic Regression": mdls(1, 2) = "0.0495": mdls(1, 3) = "~0.50": mdls(1, 4) = "0.207": mdls(1, 5) = "0.01"
    mdls(2, 1) = "EBM": mdls(2, 2) = "0.0485": mdls(2, 3) = "0.495": mdls(2, 4) = "0.207": mdls(2, 5) = "0.17"
    mdls(3, 1) = "XGBoost": mdls(3, 2) = "0.0494": mdls(3, 3) = "0.502": mdls(3, 4) = "0.207": mdls(3, 5) = "0.01"
    mClrs(1) = CLR_ACCENT: mClrs(2) = CLR_TEAL: mClrs(3) = CLR_RED
    For i = 1 To 3
        cx = 0.8 + (i - 1) * 4.1
        Call AddRoundBox(sld, cx, 1.6, 3.8, 4.2, CLR_WHITE)
        Call AddBox(sld, cx, 1.6, 3.8, 0.07, mClrs(i))
        Call AddTxt(sld, cx + 0.3, 1.85, 3.2, 0.4, mdls(i, 1), 16, mClrs(i), True, ppAlignCenter)
        Call AddTxt(sld, cx + 0.3, 2.4, 3.2, 0.25, "PR-AUC", 12, CLR_GRAY, False, ppAlignCenter)
        Call AddTxt(sld, cx + 0.3, 2.65, 3.2, 0.6, mdls(i, 2), 36, CLR_DARK, True, ppAlignCenter, FNT_HEAD)
        Call AddBox(sld, cx + 0.5, 3.4, 2.8, 0.01, CLR_LTGRAY)
        Call AddTxt(sld, cx + 0.3, 3.55, 1.6, 0.3, "ROC-AUC", 12, CLR_GRAY, False)
        Call AddTxt(sld, cx + 1.8, 3.55, 1.7, 0.3, mdls(i, 3), 14, CLR_DARK, True, ppAlignRight)
        Call AddTxt(sld, cx + 0.3, 3.95, 1.6, 0.3, "F2 (tuned)", 12, CLR_GRAY, False)
        Call AddTxt(sld, cx + 1.8, 3.95, 1.7, 0.3, mdls(i, 4), 14, CLR_DARK, True, ppAlignRight)
        Call AddTxt(sld, cx + 0.3, 4.35, 1.6, 0.3, "Threshold", 12, CLR_GRAY, False)
        Call AddTxt(sld, cx + 1.8, 4.35, 1.7, 0.3, mdls(i, 5), 14, CLR_DARK, True, ppAlignRight)
        Call AddTxt(sld, cx + 0.3, 4.85, 3.2, 0.3, "Recall @tuned = 1.00", 13, CLR_RED, True, ppAlignCenter)
    Next i
    Call AddRoundBox(sld, 0.8, 6.1, 11.7, 1.1, CLR_WHITE)
    Call AddAccentBar(sld, 0.8, 6.1, 1.1, CLR_RED)
    Call AddTxt(sld, 1.2, 6.2, 11, 0.35, "INTERPRETATION", 14, CLR_RED, True)
    Call AddTxt(sld, 1.2, 6.55, 11, 0.5, _
        "All three models converge to PR-AUC " & ChrW(8776) & " 0.05 (random baseline for 5% churn). " & _
        "At F2-optimal thresholds, all predict everyone as a churner " & ChrW(8212) & " a degenerate solution.", _
        13, CLR_DARKTEXT, False)
End Sub

' ════════════════════════════════════════════════════════════
' SLIDE: FINAL COMPARISON
' ════════════════════════════════════════════════════════════
Private Sub Slide_FinalComparison(prs As Presentation)
    Dim sld As Slide, shp As Shape, tbl As Table, r As Integer, c As Integer
    Dim td(1 To 4, 1 To 6) As String
    Set sld = prs.Slides.Add(prs.Slides.Count + 1, ppLayoutBlank)
    Call ContentSlideSetup(sld, "5.4  Final Model Comparison", "Same test set  |  Same metrics  |  Same conclusion")
    td(1, 1) = "Model": td(1, 2) = "PR-AUC": td(1, 3) = "F2": td(1, 4) = "Recall": td(1, 5) = "Precision": td(1, 6) = "Threshold"
    td(2, 1) = "Logistic Regression": td(2, 2) = "0.0495": td(2, 3) = "0.2070": td(2, 4) = "1.00": td(2, 5) = "0.0496": td(2, 6) = "0.01"
    td(3, 1) = "EBM": td(3, 2) = "0.0485": td(3, 3) = "0.2070": td(3, 4) = "1.00": td(3, 5) = "0.0496": td(3, 6) = "0.17"
    td(4, 1) = "XGBoost": td(4, 2) = "0.0494": td(4, 3) = "0.2069": td(4, 4) = "1.00": td(4, 5) = "0.0496": td(4, 6) = "0.01"
    Set shp = sld.Shapes.AddTable(4, 6, 0.8 * 72, 1.6 * 72, 11.7 * 72, 2.4 * 72)
    Set tbl = shp.Table
    For r = 1 To 4: For c = 1 To 6: Call FormatTableCell(tbl, r, c, td(r, c), 14, 4): Next c: Next r
    Call AddRoundBox(sld, 0.8, 4.4, 5.6, 2.5, CLR_WHITE)
    Call AddBox(sld, 0.8, 4.4, 0.08, 2.5, CLR_ACCENT)
    Call AddTxt(sld, 1.2, 4.6, 4.8, 0.35, "RECOMMENDED MODEL", 15, CLR_ACCENT, True)
    Call AddTxt(sld, 1.2, 5#, 4.8, 0.5, "Logistic Regression", 24, CLR_DARK, True, ppAlignLeft, FNT_HEAD)
    Call AddTxt(sld, 1.2, 5.55, 4.8, 1, _
        "Best PR-AUC (0.0495) and F2 (0.2070). Simplest, most interpretable, cheapest to deploy. " & _
        "All models are effectively random.", 13, CLR_DARKTEXT, False)
    Call AddRoundBox(sld, 6.8, 4.4, 5.7, 2.5, CLR_WHITE)
    Call AddBox(sld, 6.8, 4.4, 0.08, 2.5, CLR_RED)
    Call AddTxt(sld, 7.2, 4.6, 5, 0.35, "ROOT CAUSE", 15, CLR_RED, True)
    Call AddTxt(sld, 7.2, 5#, 5, 1.5, _
        "The bottleneck is not the model " & ChrW(8212) & " it is the feature set. " & _
        "All 13 core features show zero discriminative signal. " & _
        "No algorithm can learn patterns that do not exist in the data.", 14, CLR_DARKTEXT, False)
End Sub

' ════════════════════════════════════════════════════════════
' SLIDE: FINAL VERDICT
' ════════════════════════════════════════════════════════════
Private Sub Slide_FinalVerdict(prs As Presentation)
    Dim sld As Slide
    Set sld = prs.Slides.Add(prs.Slides.Count + 1, ppLayoutBlank)
    Call SetDarkBg(sld): Call AddBox(sld, 0, 0, 0.15, SH, CLR_RED)
    Call AddTxt(sld, 0.8, 0.8, 11.5, 0.5, "SECTIONS 1" & ChrW(8211) & "5 VERDICT", 15, CLR_ACCENT, True)
    Call AddBox(sld, 0.8, 1.4, 4, 0.04, CLR_ACCENT)
    Call AddTxt(sld, 0.8, 1.8, 11.5, 1.2, _
        "Models don't fail because of modeling." & vbCrLf & "They fail because the data has no signal.", _
        30, CLR_WHITE, True, ppAlignLeft, FNT_HEAD)
    Call AddRoundBox(sld, 0.8, 3.3, 3.7, 2.6, CLR_CARD)
    Call AddBox(sld, 0.8, 3.3, 3.7, 0.06, CLR_ACCENT)
    Call AddTxt(sld, 1.1, 3.55, 3.2, 0.35, "STATISTICAL", 14, CLR_ACCENT, True)
    Call AddTxt(sld, 1.1, 4#, 3.2, 1.6, _
        "All t-tests, Chi-square, correlation: p > 0.05" & vbCrLf & _
        "Cohen's d < 0.03 everywhere" & vbCrLf & _
        "Even engineered features: |r| < 0.01", 13, CLR_ICE, False)
    Call AddRoundBox(sld, 4.8, 3.3, 3.7, 2.6, CLR_CARD)
    Call AddBox(sld, 4.8, 3.3, 3.7, 0.06, CLR_TEAL)
    Call AddTxt(sld, 5.1, 3.55, 3.2, 0.35, "EDA", 14, CLR_TEAL, True)
    Call AddTxt(sld, 5.1, 4#, 3.2, 1.6, _
        "Perfect overlap in all distributions" & vbCrLf & _
        "All income quintiles: ~4.96%" & vbCrLf & _
        "All age bands: ~4.96%" & vbCrLf & _
        "0 high-risk countries", 13, CLR_ICE, False)
    Call AddRoundBox(sld, 8.8, 3.3, 3.7, 2.6, CLR_CARD)
    Call AddBox(sld, 8.8, 3.3, 3.7, 0.06, CLR_RED)
    Call AddTxt(sld, 9.1, 3.55, 3.2, 0.35, "MODELING", 14, CLR_RED, True)
    Call AddTxt(sld, 9.1, 4#, 3.2, 1.6, _
        "LR, EBM, XGBoost: PR-AUC " & ChrW(8776) & " 0.05" & vbCrLf & _
        "Decision Tree: AUC = 0.503" & vbCrLf & _
        "Sec 3 XGBoost: AUC = 0.499" & vbCrLf & _
        "Tuned XGBoost: AUC = 0.502", 13, CLR_ICE, False)
    Call AddRoundBox(sld, 0.8, 6.2, 11.7, 0.9, CLR_CARD)
    Call AddTxt(sld, 1.2, 6.32, 11, 0.6, _
        "We need more and better data. The production pipeline is built " & ChrW(8212) & " ready to ingest new features at scale.", _
        15, CLR_GOLD, False, ppAlignCenter)
End Sub

' ════════════════════════════════════════════════════════════
' SLIDE: PRODUCTION VISION
' ════════════════════════════════════════════════════════════
Private Sub Slide_ProductionVision(prs As Presentation)
    Dim sld As Slide, i As Integer, cy As Single
    Dim caps(1 To 5, 1 To 3) As String, capClrs(1 To 5) As Long
    Set sld = prs.Slides.Add(prs.Slides.Count + 1, ppLayoutBlank)
    Call ContentSlideSetup(sld, "Production Pipeline  " & ChrW(8212) & "  Architecture Overview", _
        "Scalable, deployable, explainable " & ChrW(8212) & " built and ready")
    caps(1, 1) = "1": caps(1, 2) = "Data Ingestion & Validation"
    caps(1, 3) = "FastAPI + Pydantic schemas. Auto-reject malformed rows with 422 errors."
    caps(2, 1) = "2": caps(2, 2) = "Batch & Single Prediction"
    caps(2, 3) = "/predict/single for real-time  |  /predict/batch with CSV upload for end-of-day runs."
    caps(3, 1) = "3": caps(3, 2) = "Explainability Layer"
    caps(3, 3) = "SHAP TreeExplainer for population-level drivers. Per-client attrition probability + top features."
    caps(4, 1) = "4": caps(4, 2) = "App Lifecycle & State"
    caps(4, 3) = "Model loaded once at startup via @asynccontextmanager. No disk I/O per request."
    caps(5, 1) = "5": caps(5, 2) = "Containerized Deployment"
    caps(5, 3) = "Docker (python:3.10-slim) + uvicorn. Identical: local, staging, production."
    capClrs(1) = CLR_ACCENT: capClrs(2) = CLR_TEAL: capClrs(3) = CLR_GOLD: capClrs(4) = CLR_NAVY: capClrs(5) = CLR_RED
    For i = 1 To 5
        cy = 1.45 + (i - 1) * 1.12
        Call AddRoundBox(sld, 0.8, cy, 11.7, 0.95, CLR_WHITE)
        Call AddBox(sld, 0.8, cy, 0.08, 0.95, capClrs(i))
        Call AddTxt(sld, 1.15, cy + 0.15, 0.6, 0.5, caps(i, 1), 22, capClrs(i), True, ppAlignCenter, FNT_HEAD)
        Call AddTxt(sld, 2#, cy + 0.08, 4, 0.4, caps(i, 2), 15, CLR_DARK, True)
        Call AddTxt(sld, 2#, cy + 0.48, 10, 0.4, caps(i, 3), 13, CLR_DARKTEXT, False)
    Next i
End Sub

' ════════════════════════════════════════════════════════════
' SLIDE: CLOSING
' ════════════════════════════════════════════════════════════
Private Sub Slide_Closing(prs As Presentation)
    Dim sld As Slide
    Set sld = prs.Slides.Add(prs.Slides.Count + 1, ppLayoutBlank)
    Call SetDarkBg(sld): Call AddBox(sld, 0, 0, 0.15, SH, CLR_ACCENT)
    Call AddTxt(sld, 0.8, 2#, 11.5, 0.8, "Thank You", 48, CLR_WHITE, True, ppAlignCenter, FNT_HEAD)
    Call AddBox(sld, 5.5, 2.9, 2.3, 0.04, CLR_ACCENT)
    Call AddTxt(sld, 0.8, 3.3, 11.5, 0.8, _
        "Credit Card Customer Attrition " & ChrW(8212) & " Predictive Modeling & Analysis", _
        18, CLR_ICE, False, ppAlignCenter)
    Call AddStat(sld, 1.5, 4.5, 3, "0.05", "Best PR-AUC (random = 0.05)", CLR_RED)
    Call AddStat(sld, 5.2, 4.5, 3, "6/6", "Sections Complete", CLR_TEAL)
    Call AddStat(sld, 8.8, 4.5, 3, "BUILT", "Production Pipeline", CLR_GOLD)
    Call AddTxt(sld, 0.8, 6.5, 11.5, 0.4, _
        "The methodology was rigorous. The data needs enrichment. The pipeline is built and ready to scale.", _
        15, CLR_ICE, False, ppAlignCenter)
End Sub

' ════════════════════════════════════════════════════════════
' SECTION 6 DIVIDER
' ════════════════════════════════════════════════════════════
Private Sub Slide_Section6_Divider(prs As Presentation)
    Call SectionDivider(prs, "06", "Production Pipeline — Built", _
        "FastAPI  |  PSI Drift Monitoring  |  Rule-Based Recommender  |  Docker")
End Sub

' ════════════════════════════════════════════════════════════
' SLIDE: 6.1 PREDICTION API ENDPOINTS
' ════════════════════════════════════════════════════════════
Private Sub Slide_APIEndpoints(prs As Presentation)
    Dim sld As Slide
    Set sld = prs.Slides.Add(prs.Slides.Count + 1, ppLayoutBlank)
    Call ContentSlideSetup(sld, "6.1  Prediction API  " & ChrW(8212) & "  Endpoints", _
        "FastAPI + Pydantic  |  Model loaded once at startup via @asynccontextmanager")

    ' /predict/single card
    Call AddRoundBox(sld, 0.8, 1.6, 5.7, 4, CLR_WHITE)
    Call AddBox(sld, 0.8, 1.6, 5.7, 0.07, CLR_ACCENT)
    Call AddTxt(sld, 1.1, 1.85, 5, 0.35, "/predict/single", 18, CLR_ACCENT, True, ppAlignLeft, FNT_HEAD)
    Call AddTxt(sld, 1.1, 2.3, 5.1, 0.3, "Input: JSON  |  11 client fields  |  Pydantic-validated", 13, CLR_GRAY, False)
    Call AddTxt(sld, 1.1, 2.75, 5.1, 2.5, _
        "attrition_probability: 0.73" & vbCrLf & _
        "attrition_flag: 1" & vbCrLf & _
        "attrition_risk_tier: ""High""" & vbCrLf & _
        "top_drivers: { Credit_Utilization: -0.42, ... }" & vbCrLf & _
        "recommended_products: [" & vbCrLf & _
        "  { product: ""Rewards Upgrade""," & vbCrLf & _
        "    reason: ""High churn risk"" }" & vbCrLf & _
        "]", 12, CLR_DARKTEXT, False)

    ' /predict/batch card
    Call AddRoundBox(sld, 6.8, 1.6, 5.7, 4, CLR_WHITE)
    Call AddBox(sld, 6.8, 1.6, 5.7, 0.07, CLR_TEAL)
    Call AddTxt(sld, 7.1, 1.85, 5, 0.35, "/predict/batch/upload", 18, CLR_TEAL, True, ppAlignLeft, FNT_HEAD)
    Call AddTxt(sld, 7.1, 2.3, 5.1, 0.3, "Input: CSV upload  |  UploadFile  |  Returns JSON array", 13, CLR_GRAY, False)
    Call AddTxt(sld, 7.1, 2.75, 5.1, 2.5, _
        "Reads CSV " & ChrW(8594) & " Pandas DataFrame" & vbCrLf & _
        "Validates all rows via schema" & vbCrLf & _
        "Runs pipeline.predict_proba()" & vbCrLf & _
        "Applies recommender per row" & vbCrLf & _
        "Returns full JSON array" & vbCrLf & _
        "Logs all predictions to" & vbCrLf & _
        "logs/predictions.csv", 12, CLR_DARKTEXT, False)

    ' Footer note
    Call AddRoundBox(sld, 0.8, 6.0, 11.7, 0.75, CLR_DARK)
    Call AddTxt(sld, 1.2, 6.1, 11, 0.5, _
        "Invalid input on any field returns HTTP 422 with the exact missing/malformed field  " & ChrW(8212) & "  zero custom validation code required.", _
        13, CLR_ICE, False, ppAlignCenter)
End Sub

' ════════════════════════════════════════════════════════════
' SLIDE: 6.2 DATA DRIFT MONITORING
' ════════════════════════════════════════════════════════════
Private Sub Slide_DriftMonitoring(prs As Presentation)
    Dim sld As Slide, i As Integer, cx As Single
    Dim tiers(1 To 3, 1 To 3) As String, tClrs(1 To 3) As Long
    Set sld = prs.Slides.Add(prs.Slides.Count + 1, ppLayoutBlank)
    Call ContentSlideSetup(sld, "6.2  Data Drift Monitoring  " & ChrW(8212) & "  PSI", _
        "Population Stability Index  |  SR 11-7 banking standard  |  Run weekly via scripts/monitor_drift.py")

    ' PSI tier cards
    tiers(1, 1) = "PSI < 0.10": tiers(1, 2) = "STABLE": tiers(1, 3) = "No action needed"
    tiers(2, 1) = "PSI < 0.20": tiers(2, 2) = "MODERATE": tiers(2, 3) = "Monitor closely"
    tiers(3, 1) = "PSI " & ChrW(8805) & " 0.20": tiers(3, 2) = "ALERT": tiers(3, 3) = "Investigate & retrain"
    tClrs(1) = CLR_GREEN: tClrs(2) = CLR_GOLD: tClrs(3) = CLR_RED
    For i = 1 To 3
        cx = 0.8 + (i - 1) * 4.05
        Call AddRoundBox(sld, cx, 1.6, 3.8, 1.8, CLR_WHITE)
        Call AddBox(sld, cx, 1.6, 3.8, 0.07, tClrs(i))
        Call AddTxt(sld, cx + 0.3, 1.85, 3.2, 0.35, tiers(i, 1), 16, CLR_DARK, True, ppAlignCenter, FNT_HEAD)
        Call AddTxt(sld, cx + 0.3, 2.3, 3.2, 0.45, tiers(i, 2), 22, tClrs(i), True, ppAlignCenter, FNT_HEAD)
        Call AddTxt(sld, cx + 0.3, 2.85, 3.2, 0.3, tiers(i, 3), 13, CLR_GRAY, False, ppAlignCenter)
    Next i

    ' How it works — flow
    Call AddTxt(sld, 0.8, 3.75, 11, 0.35, "HOW IT WORKS", 14, CLR_DARK, True)
    Call AddRoundBox(sld, 0.8, 4.2, 11.7, 1.05, CLR_WHITE)

    Dim steps(1 To 4, 1 To 2) As String
    steps(1, 1) = "1.  Serialize": steps(1, 2) = "serialize_model.py saves training_profile.json" & vbCrLf & "(feature distributions + score dist.)"
    steps(2, 1) = "2.  Log": steps(2, 2) = "Every API prediction appended to" & vbCrLf & "logs/predictions.csv with timestamp"
    steps(3, 1) = "3.  Compare": steps(3, 2) = "monitor_drift.py computes PSI" & vbCrLf & "per feature vs. training baseline"
    steps(4, 1) = "4.  Alert": steps(4, 2) = "PSI " & ChrW(8805) & " 0.20 on any feature or score" & vbCrLf & "triggers retraining flag"

    Dim i2 As Integer
    For i2 = 1 To 4
        cx = 0.85 + (i2 - 1) * 2.95
        Call AddTxt(sld, cx, 4.3, 2.6, 0.3, steps(i2, 1), 12, CLR_ACCENT, True)
        Call AddTxt(sld, cx, 4.6, 2.7, 0.55, steps(i2, 2), 11, CLR_DARKTEXT, False)
    Next i2

    ' Monitored features
    Call AddRoundBox(sld, 0.8, 5.55, 11.7, 0.65, CLR_DARK)
    Call AddTxt(sld, 1.2, 5.65, 11, 0.45, _
        "Monitored:  Age  |  Income  |  CreditLimit  |  TotalTransactions  |  TotalSpend  |  Tenure  |  attrition_probability (score drift)", _
        13, CLR_ICE, False, ppAlignCenter)
End Sub

' ════════════════════════════════════════════════════════════
' SLIDE: 6.3 RULE-BASED RECOMMENDER ENGINE
' ════════════════════════════════════════════════════════════
Private Sub Slide_RecommenderEngine(prs As Presentation)
    Dim sld As Slide, i As Integer, cy As Single
    Dim rules(1 To 5, 1 To 3) As String, rClrs(1 To 5) As Long
    Set sld = prs.Slides.Add(prs.Slides.Count + 1, ppLayoutBlank)
    Call ContentSlideSetup(sld, "6.3  Rule-Based Recommender Engine", _
        "Post-prediction layer  |  api/recommender.py  |  Zero ML  |  Zero retraining")

    ' Rule rows
    rules(1, 1) = "Attrition Risk": rules(1, 2) = "Probability " & ChrW(8805) & " 0.70": rules(1, 3) = "Rewards Upgrade  |  Retention cashback"
    rules(2, 1) = "Card Upgrade": rules(2, 2) = "Income " & ChrW(8805) & " 80K + Blue/Silver card": rules(2, 3) = "UnionBank Gold / Platinum Card"
    rules(3, 1) = "Wealth Tier": rules(3, 2) = "Income " & ChrW(8805) & " 150K": rules(3, 3) = "Wealth Management  |  Priority Banking"
    rules(4, 1) = "Loyalty": rules(4, 2) = "Tenure " & ChrW(8805) & " 48 months + prob " & ChrW(8805) & " 0.40": rules(4, 3) = "Loyalty Cashback  |  Tenure reward"
    rules(5, 1) = "New Client": rules(5, 2) = "Tenure < 12 months": rules(5, 3) = "Digital Onboarding  |  UnionBank Online"
    rClrs(1) = CLR_RED: rClrs(2) = CLR_GOLD: rClrs(3) = CLR_TEAL: rClrs(4) = CLR_ACCENT: rClrs(5) = CLR_NAVY

    For i = 1 To 5
        cy = 1.55 + (i - 1) * 0.82
        Call AddRoundBox(sld, 0.8, cy, 11.7, 0.68, CLR_WHITE)
        Call AddBox(sld, 0.8, cy, 0.07, 0.68, rClrs(i))
        Call AddTxt(sld, 1.1, cy + 0.15, 2.2, 0.35, rules(i, 1), 14, rClrs(i), True)
        Call AddTxt(sld, 3.5, cy + 0.15, 3.5, 0.35, rules(i, 2), 13, CLR_DARKTEXT, False)
        Call AddTxt(sld, 7.3, cy + 0.15, 4.8, 0.35, rules(i, 3), 13, CLR_GRAY, False)
    Next i

    ' Column headers
    Call AddTxt(sld, 1.1, 1.35, 2.2, 0.3, "Rule Layer", 11, CLR_GRAY, True)
    Call AddTxt(sld, 3.5, 1.35, 3.5, 0.3, "Trigger Condition", 11, CLR_GRAY, True)
    Call AddTxt(sld, 7.3, 1.35, 4.8, 0.3, "Recommended Products", 11, CLR_GRAY, True)

    ' Footer
    Call AddRoundBox(sld, 0.8, 6.0, 11.7, 0.75, CLR_DARK)
    Call AddTxt(sld, 1.2, 6.12, 11, 0.5, _
        "Pipeline ends at probability  " & ChrW(8594) & "  recommender reads output + raw features  " & ChrW(8594) & _
        "  model never retrained. Swap rules independently of the ML artifact.", _
        13, CLR_ICE, False, ppAlignCenter)
End Sub

' ════════════════════════════════════════════════════════════
' SLIDE: BUSINESS VALUE OF ATTRITION PREDICTION
' ════════════════════════════════════════════════════════════
Private Sub Slide_BusinessValue(prs As Presentation)
    Dim sld As Slide, i As Integer, cx As Single
    Dim rows(1 To 4, 1 To 3) As String

    Set sld = prs.Slides.Add(prs.Slides.Count + 1, ppLayoutBlank)
    Call ContentSlideSetup(sld, "Why Prediction Has Monetary Value", _
        "Early identification of at-risk clients unlocks measurable retention value  |  100K-client portfolio")

    Dim PHP As String: PHP = ChrW(8369)   ' ₱

    ' ── 3 stat cards ──────────────────────────────────────────
    Call AddRoundBox(sld, 0.8, 1.55, 3.7, 1.75, CLR_WHITE)
    Call AddBox(sld, 0.8, 1.55, 3.7, 0.07, CLR_RED)
    Call AddTxt(sld, 1.1, 1.75, 3.1, 0.35, "Cost to Acquire 1 Client", 13, CLR_GRAY, False, ppAlignCenter)
    Call AddTxt(sld, 1.1, 2.1, 3.1, 0.55, PHP & "5K " & ChrW(8211) & " " & PHP & "15K", 28, CLR_RED, True, ppAlignCenter, FNT_HEAD)
    Call AddTxt(sld, 1.1, 2.75, 3.1, 0.3, "vs. ~" & PHP & "500 to retain  (BSP, 2023)", 12, CLR_GRAY, False, ppAlignCenter)

    Call AddRoundBox(sld, 4.8, 1.55, 3.7, 1.75, CLR_WHITE)
    Call AddBox(sld, 4.8, 1.55, 3.7, 0.07, CLR_GOLD)
    Call AddTxt(sld, 5.1, 1.75, 3.1, 0.35, "Avg. Credit Card CLV (PH)", 13, CLR_GRAY, False, ppAlignCenter)
    Call AddTxt(sld, 5.1, 2.1, 3.1, 0.55, PHP & "30K " & ChrW(8211) & " " & PHP & "60K", 28, CLR_GOLD, True, ppAlignCenter, FNT_HEAD)
    Call AddTxt(sld, 5.1, 2.75, 3.1, 0.3, "3" & ChrW(8211) & "5 yr tenure " & ChrW(215) & " ~" & PHP & "9K/yr net revenue", 12, CLR_GRAY, False, ppAlignCenter)

    Call AddRoundBox(sld, 8.8, 1.55, 3.7, 1.75, CLR_WHITE)
    Call AddBox(sld, 8.8, 1.55, 3.7, 0.07, CLR_ACCENT)
    Call AddTxt(sld, 9.1, 1.75, 3.1, 0.35, "Retention " & ChrW(8593) & "5%  " & ChrW(8594) & "  Profit", 13, CLR_GRAY, False, ppAlignCenter)
    Call AddTxt(sld, 9.1, 2.1, 3.1, 0.55, "+25 " & ChrW(8211) & " 55%", 28, CLR_ACCENT, True, ppAlignCenter, FNT_HEAD)
    Call AddTxt(sld, 9.1, 2.75, 3.1, 0.3, "Bain & Company (2013)", 12, CLR_GRAY, False, ppAlignCenter)

    ' ── Expected Value Table ──────────────────────────────────
    Call AddTxt(sld, 0.8, 3.55, 9, 0.35, "EXPECTED VALUE FRAMEWORK  " & ChrW(8212) & "  100K Portfolio  |  5% Annual Churn  |  Avg CLV " & PHP & "36,000", 14, CLR_DARK, True)
    rows(1, 1) = "At-risk clients (5% of 100K)": rows(1, 2) = "5,000 clients": rows(1, 3) = ""
    rows(2, 1) = "Revenue at risk (5,000 " & ChrW(215) & " " & PHP & "36,000 avg CLV)": rows(2, 2) = PHP & "180,000,000": rows(2, 3) = "annually"
    rows(3, 1) = "Targeted intervention (top decile, 40% success rate)": rows(3, 2) = "~200 retained": rows(3, 3) = "500 contacted " & ChrW(215) & " 40%"
    rows(4, 1) = "Net revenue preserved (200 " & ChrW(215) & " " & PHP & "36K CLV " & ChrW(8722) & " " & PHP & "250K cost)": rows(4, 2) = "~" & PHP & "6,950,000": rows(4, 3) = "per intervention cycle"

    For i = 1 To 4
        Dim cy As Single: cy = 4.05 + (i - 1) * 0.48
        Dim rClr As Long
        If i = 4 Then rClr = CLR_GREEN Else rClr = CLR_OFFWHITE
        Call AddRoundBox(sld, 0.8, cy, 11.7, 0.42, rClr)
        Call AddTxt(sld, 1.1, cy + 0.06, 7.2, 0.3, rows(i, 1), 13, CLR_DARKTEXT, IIf(i = 4, True, False))
        Call AddTxt(sld, 8.5, cy + 0.06, 2.2, 0.3, rows(i, 2), 14, IIf(i = 4, CLR_DARK, CLR_DARK), True, ppAlignRight)
        Call AddTxt(sld, 10.9, cy + 0.06, 1.5, 0.3, rows(i, 3), 11, CLR_GRAY, False)
    Next i

    ' ── Closing statement ─────────────────────────────────────
    Call AddRoundBox(sld, 0.8, 6.1, 11.7, 0.65, CLR_DARK)
    Call AddTxt(sld, 1.2, 6.2, 11, 0.45, _
        "A model that identifies the right 500 clients out of 100,000 — before they leave — is worth " & PHP & "6.95M in preserved revenue per cycle.", _
        13, CLR_ICE, False, ppAlignCenter)

    ' ── Sources ──────────────────────────────────────────────
    Dim srcTxt As String
    srcTxt = "Sources:  (1) BSP, " & Chr(34) & "Credit Card Industry Report & Financial Stability Report," & Chr(34) & " 2023  " & ChrW(183) & _
             "  (2) Bangko Sentral ng Pilipinas, " & Chr(34) & "Consumer Finance Survey," & Chr(34) & " 2022  " & ChrW(183) & _
             "  (3) Bain & Company, " & Chr(34) & "Customer Loyalty in Retail Banking," & Chr(34) & " 2013  " & ChrW(183) & _
             "  (4) McKinsey & Company, " & Chr(34) & "Remaking the bank for an ecosystem world," & Chr(34) & " 2017  " & ChrW(183) & _
             "  (5) Reichheld & Sasser, " & Chr(34) & "Zero Defections: Quality Comes to Services," & Chr(34) & " Harvard Business Review, 1990"
    Call AddTxt(sld, 0.8, 7.25, 11.7, 0.3, srcTxt, 8.5, CLR_GRAY, False, ppAlignLeft)
End Sub
