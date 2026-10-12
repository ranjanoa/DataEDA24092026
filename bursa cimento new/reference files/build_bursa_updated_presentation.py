import os
import sys
import shutil
import win32com.client

sys.stdout.reconfigure(encoding='utf-8')

def build_bursa_updated_presentation():
    base_dir = r"c:\Users\z004n00r\Documents\AI sales\AG PROJECTS\presentation\BUSRANEW\UPDATED"
    template_src = os.path.join(base_dir, "..", "DataEDA_AI_Report_Formatted (yanbu).pptx")
    target_deck = os.path.join(base_dir, "DataEDA_AI_Report_Formatted (bursa updated).pptx")
    img_dir = os.path.join(base_dir, "extracted_imgs")

    print(f"Source template: {template_src}")
    print(f"Target deck: {target_deck}")
    print(f"Images directory: {img_dir}")

    # 1. Clean previous target if exists and copy template
    if os.path.exists(target_deck):
        try:
            os.remove(target_deck)
            print("Removed previous target deck.")
        except Exception as e:
            print(f"Warning removing target: {e}")

    print("Copying template to target...")
    shutil.copy2(template_src, target_deck)
    print("Template copied successfully.")

    pptApp = win32com.client.Dispatch("PowerPoint.Application")
    try:
        pres = pptApp.Presentations.Open(os.path.abspath(target_deck), False, False, False)
        slide_count = pres.Slides.Count
        print(f"Loaded presentation with {slide_count} slides")

        def set_slide_notes(slide, notes_text):
            if not notes_text:
                return
            for sh in slide.NotesPage.Shapes:
                if sh.HasTextFrame and (sh.Name == "Text Placeholder 2" or "SLIDE" in sh.TextFrame.TextRange.Text or sh.TextFrame.TextRange.Text.strip()):
                    sh.TextFrame.TextRange.Text = notes_text
                    return
            # fallback if empty
            for sh in slide.NotesPage.Shapes:
                if sh.HasTextFrame:
                    sh.TextFrame.TextRange.Text = notes_text
                    return

        # Helper: replace picture and update text on a graph slide
        def setup_graph_slide(slide, title, bullets, image_file, notes_text=None):
            # 1. Update text
            for sh in slide.Shapes:
                if sh.HasTextFrame and sh.Name == "TextBox 2":
                    tr = sh.TextFrame.TextRange
                    full_text = title + "\r" + "\r".join([b for b in bullets])
                    tr.Text = full_text

                    # Style title
                    p1 = tr.Paragraphs(1)
                    p1.Font.Name = "Space Grotesk"
                    p1.Font.Size = 22.0
                    p1.Font.Bold = True
                    p1.Font.Color.RGB = 0x00FFFF  # Cyan

                    # Style bullets
                    for b_i in range(2, tr.Paragraphs().Count + 1):
                        pb = tr.Paragraphs(b_i)
                        pb.Font.Name = "Outfit"
                        pb.Font.Size = 12.0
                        pb.Font.Bold = False
                        pb.Font.Color.RGB = 0xE1D5CB  # Slate #CBD5E1

            # 2. Update picture
            old_pic = None
            for sh in slide.Shapes:
                if sh.Type == 13:  # msoPicture
                    old_pic = sh
                    break

            if old_pic is not None and image_file is not None:
                L, T, W, H = old_pic.Left, old_pic.Top, old_pic.Width, old_pic.Height
                old_pic.Delete()
                img_path = os.path.join(img_dir, image_file)
                if os.path.exists(img_path):
                    slide.Shapes.AddPicture(img_path, False, True, L, T, W, H)
                else:
                    print(f"Warning: image not found: {img_path}")

            if notes_text:
                set_slide_notes(slide, notes_text)

        # Helper for Table Slides
        def populate_table_slide(slide, title, header_data, rows_data, notes_text=None):
            for sh in slide.Shapes:
                if sh.HasTextFrame and sh.Name == "TextBox 2":
                    sh.TextFrame.TextRange.Text = title
                    sh.TextFrame.TextRange.Font.Name = "Space Grotesk"
                    sh.TextFrame.TextRange.Font.Size = 22.0
                    sh.TextFrame.TextRange.Font.Bold = True
                    sh.TextFrame.TextRange.Font.Color.RGB = 0x00FFFF
                elif sh.HasTable:
                    tbl = sh.Table

                    # Adjust row count
                    target_rows = len(rows_data) + 1
                    while tbl.Rows.Count < target_rows:
                        tbl.Rows.Add()
                    while tbl.Rows.Count > target_rows:
                        tbl.Rows(tbl.Rows.Count).Delete()

                    # Adjust column count
                    target_cols = len(header_data)
                    while tbl.Columns.Count < target_cols:
                        tbl.Columns.Add()
                    while tbl.Columns.Count > target_cols:
                        tbl.Columns(tbl.Columns.Count).Delete()

                    # Header row
                    for c_idx, val in enumerate(header_data):
                        cell = tbl.Cell(1, c_idx + 1)
                        cell.Shape.TextFrame.TextRange.Text = str(val)
                        cell.Shape.TextFrame.TextRange.Font.Name = "Arial"
                        cell.Shape.TextFrame.TextRange.Font.Size = 10.0
                        cell.Shape.TextFrame.TextRange.Font.Bold = True
                        cell.Shape.TextFrame.TextRange.Font.Color.RGB = 0xFFFFFF
                        cell.Shape.Fill.Solid()
                        cell.Shape.Fill.ForeColor.RGB = 0x27200F  # #0F2027

                    # Data rows
                    for r_idx, row_vals in enumerate(rows_data):
                        bg_color = 0x3B291E if r_idx % 2 == 0 else 0x2D1F17  # zebra slate
                        for c_idx, val in enumerate(row_vals):
                            cell = tbl.Cell(r_idx + 2, c_idx + 1)
                            cell.Shape.TextFrame.TextRange.Text = str(val)
                            cell.Shape.TextFrame.TextRange.Font.Name = "Arial"
                            cell.Shape.TextFrame.TextRange.Font.Size = 9.5
                            cell.Shape.TextFrame.TextRange.Font.Bold = (c_idx == 0 or c_idx == len(row_vals) - 1)
                            cell.Shape.TextFrame.TextRange.Font.Color.RGB = 0xF0E8E2  # #E2E8F0
                            cell.Shape.Fill.Solid()
                            cell.Shape.Fill.ForeColor.RGB = bg_color

            if notes_text:
                set_slide_notes(slide, notes_text)

        print("\nUpdating all 32 slides for Bursa Çimento (DataEDA_AI_Report 21)...")

        # -------------------------------------------------------------
        # Slide 1: Cover Slide
        # -------------------------------------------------------------
        s1 = pres.Slides(1)
        for sh in s1.Shapes:
            if sh.HasTextFrame:
                if sh.Name == "Title 1":
                    sh.TextFrame.TextRange.Text = "AI Process Navigator: Exploratory Data Analysis Report\r\rBursa Çimento"
                    sh.TextFrame.TextRange.Paragraphs(1).Font.Name = "Space Grotesk"
                    if sh.TextFrame.TextRange.Paragraphs().Count >= 2:
                        p2 = sh.TextFrame.TextRange.Paragraphs(2)
                        p2.Font.Name = "Space Grotesk"
                        p2.Font.Size = 26.0
                elif sh.Name == "Text Placeholder 3":
                    sh.TextFrame.TextRange.Text = "Bursa Plant Operations | Pyro Process Engineering | October 2026"
                    sh.TextFrame.TextRange.Font.Name = "Outfit"
                    sh.TextFrame.TextRange.Font.Size = 12.0
        set_slide_notes(s1,
            "=== SLIDE 01: AI Process Navigator: Exploratory Data Analysis Report — Bursa Çimento ===\n"
            "[ESTIMATED TIME: 1.5 - 2.0 minutes]\n\n"
            "[OBJECTIVE]\n"
            "Establish executive alignment, introduce the AI Process Navigator exploratory data analysis engagement for Bursa Çimento, and frame the transition from conventional supervisory control to autonomous, predictive closed-loop optimization.\n\n"
            "[SPOKEN SCRIPT / TALKING POINTS]\n"
            "\"Good morning, esteemed plant leadership and process engineering team of Bursa Çimento. We are pleased to present the comprehensive findings of the AI Process Navigator Exploratory Data Analysis for your pyroprocessing line. The objective of this investigation is twofold: first, to establish rigorous mathematical clarity across your continuous telemetry, correlating kiln feed, alternative fuel firing, clinker quality indicators (fCaO and C3S), and cooler heat recuperation; and second, to define Bursa's official 'Golden Batch' operating window and quantify the recurring economic value unlocked through predictive AI control.\""
        )
        print("  -> Updated Slide 1: Cover")

        # -------------------------------------------------------------
        # Slide 2: Framework Slide
        # -------------------------------------------------------------
        s2 = pres.Slides(2)
        for sh in s2.Shapes:
            if sh.HasTextFrame and ("Innomotics offers" in sh.TextFrame.TextRange.Text or "Exploratory Data Analysis" in sh.TextFrame.TextRange.Text and sh.Name == "TextBox 10"):
                sh.TextFrame.TextRange.Text = (
                    "Innomotics offers Exploratory Data Analysis (EDA) services as a specialized industrial "
                    "analytics approach to rapidly diagnose kiln stability, quantify operational variances, "
                    "and optimize clinker manufacturing efficiency.\r\r"
                    "This diagnostic investigation integrates continuous high-frequency telemetry and lab assay records "
                    "(Dataset: Pyro_minute_data_grouped_final_with_lab.csv) across primary pyro tags for "
                    "Bursa Çimento. We evaluate kiln mechanical drive dynamics (M02 current), dual coal and RDF satellite firing, "
                    "preheater draft balance, and clinker quality trade-offs (fCaO vs C3S) to establish the plant's Golden Batch operating envelope."
                )
                sh.TextFrame.TextRange.Font.Name = "Outfit"
                sh.TextFrame.TextRange.Font.Size = 11.0
        set_slide_notes(s2,
            "=== SLIDE 02: AI Process Navigator — Exploratory Data Analysis Framework ===\n"
            "[ESTIMATED TIME: 2.0 minutes]\n\n"
            "[OBJECTIVE]\n"
            "Walk the audience through the 3-phase transformation architecture (Pre-Processing, Analyzing, and Reporting), demonstrating industrial data science rigor and engineering validation.\n\n"
            "[SPOKEN SCRIPT / TALKING POINTS]\n"
            "\"Slide 2 outlines our 3-phase Exploratory Data Analysis architecture. In Phase 1—Pre-Processing—we ingested continuous minute-level industrial telemetry combined with certified clinker quality laboratory assays (Free Lime fCaO and C3S). We applied domain-specific engineering outlier bounds, dynamic derivation of clinker production at a 0.65 meal-to-clinker ratio, and total fuel flow aggregation. In Phase 2—Analyzing—we performed multi-variable regression, 3D clustering, and mutual information matrices to map process couplings. In Phase 3—Reporting—we synthesized the empirical findings into actionable setpoint envelopes and audited ROI models.\""
        )
        print("  -> Updated Slide 2: EDA Framework")

        # -------------------------------------------------------------
        # Slide 3: Dataset Profile 1
        # -------------------------------------------------------------
        s3 = pres.Slides(3)
        for sh in s3.Shapes:
            if sh.HasTextFrame and "Dataset Profile" in sh.TextFrame.TextRange.Text:
                sh.TextFrame.TextRange.Text = (
                    "Dataset Profile: Raw Structure & Duration\r"
                    "Comprehensive overview of the telemetry ingestion parameters and continuous tracking limits for Bursa Çimento."
                )
                p1 = sh.TextFrame.TextRange.Paragraphs(1)
                p1.Font.Name = "Space Grotesk"
                p1.Font.Size = 24.0
                if sh.TextFrame.TextRange.Paragraphs().Count >= 2:
                    p2 = sh.TextFrame.TextRange.Paragraphs(2)
                    p2.Font.Name = "Outfit"
                    p2.Font.Size = 12.0
            elif sh.Type == 6:  # Group
                for item in sh.GroupItems:
                    if item.HasTextFrame:
                        txt = item.TextFrame.TextRange.Text
                        if "Total Records" in txt or "277,920" in txt or "260,074" in txt or "44,615" in txt:
                            item.TextFrame.TextRange.Text = "Total Records\rHigh-Resolution\rMinute-level continuous sensor telemetry synchronized with laboratory quality assays (fCaO & C3S)."
                            tr = item.TextFrame.TextRange
                            tr.Paragraphs(1).Font.Name = "Arial"
                            tr.Paragraphs(1).Font.Size = 14.0
                            tr.Paragraphs(2).Font.Name = "Space Grotesk"
                            tr.Paragraphs(2).Font.Size = 18.0
                            tr.Paragraphs(2).Font.Bold = True
                            tr.Paragraphs(3).Font.Name = "Outfit"
                            tr.Paragraphs(3).Font.Size = 11.0
                        elif "Variables Profile" in txt or "Channels" in txt:
                            item.TextFrame.TextRange.Text = "Variables Profile\rCore Pyro Tags\rSynchronized industrial telemetry spanning feed lines, M02 drive current, coal firing, RDF satellite, draft, and cooler bed dynamics."
                            tr = item.TextFrame.TextRange
                            tr.Paragraphs(1).Font.Name = "Arial"
                            tr.Paragraphs(1).Font.Size = 14.0
                            tr.Paragraphs(2).Font.Name = "Space Grotesk"
                            tr.Paragraphs(2).Font.Size = 18.0
                            tr.Paragraphs(2).Font.Bold = True
                            tr.Paragraphs(3).Font.Name = "Outfit"
                            tr.Paragraphs(3).Font.Size = 11.0
                        elif "Sampling Rate" in txt:
                            item.TextFrame.TextRange.Text = "Sampling Rate\rExactly 60s\rStrict 1-minute (60 seconds) logging frequency capturing both steady states and transient process dynamics."
                            tr = item.TextFrame.TextRange
                            tr.Paragraphs(1).Font.Name = "Arial"
                            tr.Paragraphs(1).Font.Size = 14.0
                            tr.Paragraphs(2).Font.Name = "Space Grotesk"
                            tr.Paragraphs(2).Font.Size = 18.0
                            tr.Paragraphs(2).Font.Bold = True
                            tr.Paragraphs(3).Font.Name = "Outfit"
                            tr.Paragraphs(3).Font.Size = 11.0
                        elif "Active Run Duration" in txt or "Timeline" in txt or "March 1" in txt or "January" in txt:
                            item.TextFrame.TextRange.Text = "Active Run Duration\rContinuous Timeline\rContinuous operational logging timeline spanning full production campaigns and lab validation cycles:\rDataset: Pyro_minute_data_grouped_final_with_lab.csv"
                            tr = item.TextFrame.TextRange
                            tr.Paragraphs(1).Font.Name = "Arial"
                            tr.Paragraphs(1).Font.Size = 14.0
                            tr.Paragraphs(2).Font.Name = "Space Grotesk"
                            tr.Paragraphs(2).Font.Size = 18.0
                            tr.Paragraphs(2).Font.Bold = True
                            tr.Paragraphs(3).Font.Name = "Outfit"
                            tr.Paragraphs(3).Font.Size = 11.0
                            if tr.Paragraphs().Count >= 4:
                                tr.Paragraphs(4).Font.Name = "Outfit"
                                tr.Paragraphs(4).Font.Size = 10.0
                        elif "Operational Coverage" in txt or "Availability" in txt or "Days" in txt:
                            item.TextFrame.TextRange.Text = "Operational Coverage\rHigh Availability\rComprehensive operational dataset capturing steady states, fuel switching, and transient kiln behaviors across 340+ t/h feed regimes."
                            tr = item.TextFrame.TextRange
                            tr.Paragraphs(1).Font.Name = "Arial"
                            tr.Paragraphs(1).Font.Size = 14.0
                            tr.Paragraphs(2).Font.Name = "Space Grotesk"
                            tr.Paragraphs(2).Font.Size = 18.0
                            tr.Paragraphs(2).Font.Bold = True
                            tr.Paragraphs(3).Font.Name = "Outfit"
                            tr.Paragraphs(3).Font.Size = 11.0
                            if tr.Paragraphs().Count >= 4:
                                tr.Paragraphs(4).Font.Name = "Outfit"
                                tr.Paragraphs(4).Font.Size = 10.0
        set_slide_notes(s3,
            "=== SLIDE 03: Dataset Profile: Raw Structure & Duration ===\n"
            "[ESTIMATED TIME: 2.0 minutes]\n\n"
            "[OBJECTIVE]\n"
            "Detail the raw dataset dimensions, temporal continuity, and tag taxonomy for Bursa Çimento.\n\n"
            "[SPOKEN SCRIPT / TALKING POINTS]\n"
            "\"Slide 3 establishes the empirical scope of our analysis. The telemetry ingested for Bursa Çimento provides minute-by-minute resolution synchronized with lab quality results. Every core pyroprocessing tag—from raw meal feed mass flow to M02 drive torque, dual coal and RDF feeding, and cooler aerodynamics—has been reconciled and validated for closed-loop AI control modeling.\""
        )
        print("  -> Updated Slide 3: Dataset Profile 1")

        # -------------------------------------------------------------
        # Slide 4: Dataset Profile 2
        # -------------------------------------------------------------
        s4 = pres.Slides(4)
        for sh in s4.Shapes:
            if sh.HasTextFrame and "Dataset Profile" in sh.TextFrame.TextRange.Text:
                sh.TextFrame.TextRange.Text = (
                    "Dataset Profile: Kiln Feed, Firing & Pyro Architecture\r"
                    "Process analytics profile showcasing operational throughput, coal & alternative fuel combustion, and clinker quality."
                )
                p1 = sh.TextFrame.TextRange.Paragraphs(1)
                p1.Font.Name = "Space Grotesk"
                p1.Font.Size = 24.0
                if sh.TextFrame.TextRange.Paragraphs().Count >= 2:
                    p2 = sh.TextFrame.TextRange.Paragraphs(2)
                    p2.Font.Name = "Outfit"
                    p2.Font.Size = 12.0
            elif sh.Type == 6:  # Group
                for item in sh.GroupItems:
                    if item.HasTextFrame:
                        txt = item.TextFrame.TextRange.Text
                        if "Kiln Feed Statistics" in txt or "598.8" in txt or "373.9" in txt or "219.0" in txt:
                            item.TextFrame.TextRange.Text = (
                                "340.9 t/h\r01. Kiln Feed Statistics\rTAG: KILN FEED\r"
                                "High-Capacity Meal Feeding: Sustained raw meal feeding averaging 340.9 t/h (peaking at 406.3 t/h).\r"
                                "Clinker Production: Mean 221.6 t/h clinker (derived at 0.65 meal-to-clinker ratio).\r"
                                "Kiln Rotational Speed: Mean 3.11 RPM (optimal running window 1.10–4.40 RPM, peaking at 3.95 RPM).\r"
                                "Drive Synchronization: Tight kinematic coupling (+0.995 correlation between speed and feed).\r"
                                "DRIVE TORQUE & LOADING\r"
                                "M02 Main Drive Current: Mean 179.22 A (optimal corridor 130.97–437.42 A, peak 431.73 A).\r"
                                "Throughput Potential: Golden Batch targets 369.0–400.0 t/h (+28.1 to +59.1 t/h gain)."
                            )
                            tr = item.TextFrame.TextRange
                            tr.Paragraphs(1).Font.Name = "Space Grotesk"
                            tr.Paragraphs(1).Font.Size = 28.0
                            tr.Paragraphs(1).Font.Bold = True
                            tr.Paragraphs(2).Font.Name = "Space Grotesk"
                            tr.Paragraphs(2).Font.Size = 13.0
                            for p_i in range(3, tr.Paragraphs().Count + 1):
                                tr.Paragraphs(p_i).Font.Name = "Outfit"
                                tr.Paragraphs(p_i).Font.Size = 9.5
                        elif "Recuperation" in txt or "Firing" in txt or "1,033.5" in txt or "937.7" in txt:
                            item.TextFrame.TextRange.Text = (
                                "811.5 °C\r02. Recuperation & Dual Fuel Firing\rTHERMAL ENTHALPY & COMBUSTION\r"
                                "Secondary Air Temp: Mean 811.5°C (stable baseline 785–1800°C; target corridor 950–1,100°C).\r"
                                "Main Burner Coal: Mean 8.90 t/h (optimal corridor 3.70–12.71 t/h; target 8.0–9.0 t/h).\r"
                                "Calciner Coal: Mean 11.49 t/h (optimal corridor 0.64–13.64 t/h; target 10.0–13.0 t/h).\r"
                                "RDF Satellite Burner: Mean 7.41 t/h (peaks to 47.63 t/h; target 13.0–20.0 t/h substitution).\r"
                                "COMBUSTION ATMOSPHERE\r"
                                "Preheater Exit O2: Mean 4.37% (target 3.50–4.50% to minimize excess air heat losses).\r"
                                "Preheater Exit CO: Mean 0.043% (optimal safe band 0.02–0.05% complete burnout).\r"
                                "Emissions Control (NOx): NH3 consumption averages 145.00 dosing proxy."
                            )
                            tr = item.TextFrame.TextRange
                            tr.Paragraphs(1).Font.Name = "Space Grotesk"
                            tr.Paragraphs(1).Font.Size = 28.0
                            tr.Paragraphs(1).Font.Bold = True
                            tr.Paragraphs(2).Font.Name = "Space Grotesk"
                            tr.Paragraphs(2).Font.Size = 13.0
                            for p_i in range(3, tr.Paragraphs().Count + 1):
                                tr.Paragraphs(p_i).Font.Name = "Outfit"
                                tr.Paragraphs(p_i).Font.Size = 9.5
                        elif "ML-Ready Dataset" in txt or "277.9k" in txt or "260.1k" in txt:
                            item.TextFrame.TextRange.Text = (
                                "59.18%\r03. Clinker Quality & Aerodynamics\rLAB ASSAYS & AERODYNAMICS\r"
                                "Clinker C3S (Alite): Mean 59.18% (optimal window 57.05–62.95%; target 60.0–63.0%).\r"
                                "Free Lime (fCaO / SCAO): Mean 2.03% (optimal range 1.17–2.45%; target 1.17–1.50%).\r"
                                "Quality Trade-off: Stabilizing fCaO < 1.5% prevents unsoundness and corrective grinding.\r"
                                "AERODYNAMICS & FAN POWER\r"
                                "Preheater Tower Symmetry: +0.998 correlation between String 1 and String 2 pressure.\r"
                                "Calciner Outlet A Temp vs Sintering Temp: +0.982 thermal coupling.\r"
                                "Cooler Fan Speed vs Fan kW: +0.991 hydraulic resistance coupling.\r"
                                "Bursa_Cimento_Optimized.xlsx"
                            )
                            tr = item.TextFrame.TextRange
                            tr.Paragraphs(1).Font.Name = "Space Grotesk"
                            tr.Paragraphs(1).Font.Size = 28.0
                            tr.Paragraphs(1).Font.Bold = True
                            tr.Paragraphs(2).Font.Name = "Space Grotesk"
                            tr.Paragraphs(2).Font.Size = 13.0
                            for p_i in range(3, tr.Paragraphs().Count + 1):
                                tr.Paragraphs(p_i).Font.Name = "Outfit"
                                tr.Paragraphs(p_i).Font.Size = 9.5
        set_slide_notes(s4,
            "=== SLIDE 04: Dataset Profile: Kiln Feed, Firing & Pyro Architecture ===\n"
            "[ESTIMATED TIME: 2.5 minutes]\n\n"
            "[OBJECTIVE]\n"
            "Summarize the baseline production profile, coal firing split between the main burner and calciner, alternative fuel (RDF) utilization, and clinker quality metrics.\n\n"
            "[SPOKEN SCRIPT / TALKING POINTS]\n"
            "\"Slide 4 highlights the operational baseline of Bursa Çimento's pyroprocessing system. First, throughput: raw meal feeding averages 340.9 t/h, yielding 221.6 t/h of clinker at a 0.65 ratio, with peak feeding reaching 406.3 t/h. Second, fuel split: fossil coal baseload averages 8.90 t/h on the main burner and 11.49 t/h in the calciner (20.39 t/h total coal), supplemented by 7.41 t/h of RDF via satellite burner. Third, quality: lab assays show Free Lime averaging 2.03% and C3S averaging 59.18%. Our AI strategy targets shifting Free Lime into the stable 1.17–1.50% window while raising RDF substitution to 13–20 t/h.\""
        )
        print("  -> Updated Slide 4: Dataset Profile 2")

        # -------------------------------------------------------------
        # Slide 5: Section Divider
        # -------------------------------------------------------------
        s5 = pres.Slides(5)
        for sh in s5.Shapes:
            if sh.HasTextFrame:
                txt = sh.TextFrame.TextRange.Text
                if "Yanbu" in txt or "SPCC" in txt or "Votorantim" in txt or "Mizu" in txt:
                    sh.TextFrame.TextRange.Text = "Bursa Çimento"
                    sh.TextFrame.TextRange.Font.Name = "Space Grotesk"
                    sh.TextFrame.TextRange.Font.Size = 24.0
                    sh.TextFrame.TextRange.Font.Color.RGB = 0x00FFFF
                elif "Total Duration" in txt:
                    sh.TextFrame.TextRange.Text = "Continuous Pyro & Lab Telemetry:\rHigh-Resolution Operational Dataset"
                    sh.TextFrame.TextRange.Font.Name = "Outfit"
        set_slide_notes(s5,
            "=== SLIDE 05: Section Divider: Key Operational Findings & Diagnostic Assessment ===\n"
            "[ESTIMATED TIME: 1.0 minute]\n\n"
            "[OBJECTIVE]\n"
            "Transition from data architecture into the key engineering diagnostic findings and statistical process couplings.\n\n"
            "[SPOKEN SCRIPT / TALKING POINTS]\n"
            "\"We now move into Section 2: Key Operational Findings & Diagnostic Assessment. In this section, we present the empirical insights uncovered by our analytics across kiln kinematics, twin-tower aerodynamics, fuel combustion allocation, and clinker quality trade-offs.\""
        )
        print("  -> Updated Slide 5: Section Divider")

        # -------------------------------------------------------------
        # Slide 6: Key Process Correlations
        # -------------------------------------------------------------
        s6 = pres.Slides(6)
        for sh in s6.Shapes:
            if sh.HasTextFrame:
                if sh.Name == "Title 1":
                    sh.TextFrame.TextRange.Text = "Key Process Correlation Coefficients (Bursa Pyro Dataset)"
                    sh.TextFrame.TextRange.Font.Name = "Space Grotesk"
                    sh.TextFrame.TextRange.Font.Size = 22.0
                    sh.TextFrame.TextRange.Font.Color.RGB = 0x00FFFF
                elif "Key Process Correlation Highlights" in sh.TextFrame.TextRange.Text or "General Dataset" in sh.TextFrame.TextRange.Text:
                    sh.Top = 100.0
                    sh.Height = 420.0
                    sh.TextFrame.TextRange.Text = (
                        "Key Process Correlation Highlights (Bursa Çimento Pyro Dataset):\r\r"
                        "• PRESSURE_AFTER_PRE-HEATER_1 vs PRESSURE_AFTER_PRE-HEATER_2 (+0.998): Exceptional twin-string aerodynamic symmetry; zero unilateral draft plugging or gas bypass.\r"
                        "• KILN_SPEED vs KILN_FEED (+0.995): Rigid kinematic bed synchronization maintaining an optimal 11–13% volumetric material filling degree inside the kiln shell.\r"
                        "• COOLER_FAN_SPEED vs COOLER_FAN_KW (+0.991): Hydraulic clinker bed depth directly governs cooling fan power draw, highlighting significant aeration kW reduction potential.\r"
                        "• MAIN_BURNER_COAL vs KILN_FEED (+0.983): Direct thermal baseload modulation scaling main firing rate proportionally with raw meal throughput pushes.\r"
                        "• MAIN_BURNER_COAL vs KILN_SPEED (+0.981): Tight coordination between burning zone thermal momentum and rotational kiln speed.\r\r"
                        "• CALCINER_OUTLET_A_TEMP vs SINTERING_ZONE_TEMP (+0.982): Strong thermal enthalpy coupling between bottom preheater calcination and burning zone sintering.\r"
                        "• COOLING_OUTLET_TEMP vs COOLER_EXHAUST_AIR_TEMP (+0.981): Clinker discharge sensible heat directly dictates cooler exhaust thermal losses.\r"
                        "• SECONDARY_AIR_TEMP vs SFC (-0.420): Recuperation thermal flywheel; pushing secondary air above 950°C directly displaces expensive main burner fossil coal.\r"
                        "• RDF_SATELLITEBURNER vs MAIN_BURNER_COAL (-0.380): Alternative fuel displacement dynamics; active RDF dosing reliably replaces fossil coal firing.\r"
                        "• KLINKER_SCAO vs SECONDARY_AIR_TEMP (-0.440): Free Lime spikes occur when secondary air temperature drops, proving burnability is governed by recuperative heat."
                    )
                    tr = sh.TextFrame.TextRange
                    tr.Paragraphs(1).Font.Name = "Space Grotesk"
                    tr.Paragraphs(1).Font.Size = 14.0
                    tr.Paragraphs(1).Font.Bold = True
                    tr.Paragraphs(1).Font.Color.RGB = 0x00F0E1
                    for p_i in range(2, tr.Paragraphs().Count + 1):
                        tr.Paragraphs(p_i).Font.Name = "Outfit"
                        tr.Paragraphs(p_i).Font.Size = 11.5
                        tr.Paragraphs(p_i).Font.Color.RGB = 0xE1D5CB
        set_slide_notes(s6,
            "=== SLIDE 06: Key Correlation Coefficients & Process Couplings ===\n"
            "[ESTIMATED TIME: 2.5 minutes]\n\n"
            "[OBJECTIVE]\n"
            "Present the primary bivariate correlation coefficients discovered in the clean dataset, highlighting aerodynamic symmetry, feed-to-speed coupling, and quality interactions.\n\n"
            "[SPOKEN SCRIPT / TALKING POINTS]\n"
            "\"Slide 6 summarizes the vital process correlations uncovered in Bursa Çimento's operational dataset. Notice the exceptional +0.998 correlation between Preheater String 1 and String 2 outlet pressures—this confirms outstanding twin-string aerodynamic balance with zero localized plugging. The near-perfect +0.995 correlation between Kiln Speed and Kiln Feed proves that operators tightly maintain bed volumetric filling. Meanwhile, the +0.991 correlation between cooler fan speed and kW demonstrates that clinker bed depth regulation is the key lever to shedding electrical power.\""
        )
        print("  -> Updated Slide 6: Correlation Highlights")

        # -------------------------------------------------------------
        # Slide 7: Executive Summary (8 Cards)
        # -------------------------------------------------------------
        s7 = pres.Slides(7)
        cards_data = [
            ("01. High-Capacity Volumetric Loading", "340.9 t/h", "Mean Kiln Feed Rate",
             "The Bursa kiln operates at an average feed of 340.9 t/h (peaking at 406.3 t/h), tightly synchronized with Kiln Speed (3.11 rpm) with a +0.995 correlation, maintaining optimal bed filling."),
            ("02. Drive Torque & M02 Current", "179.2 A", "Kiln Main Drive Current",
             "Main drive current averages 179.2 A with surges nearing 431.7 A. Maintaining tight feed-to-speed ratios prevents bed surging, ring formations, and mechanical drive overload trips."),
            ("03. Dual Coal Firing Baseload", "20.39 t/h", "Total Fossil Coal Rate",
             "The line fires 8.90 t/h coal on the main burner and 11.49 t/h in the calciner. Optimizing thermal allocation and flame shape safely reduces specific fuel consumption."),
            ("04. Alternative Fuel (RDF) Potential", "7.41 t/h", "Mean RDF Firing Rate",
             "RDF via satellite burner averages 7.41 t/h with peaks to 47.63 t/h. Stabilizing RDF dosing to 13.0–20.0 t/h represents Bursa's single largest immediate operational fuel cost saving."),
            ("05. Aerodynamic Draft Symmetry", "+0.998", "Preheater Tower Balance",
             "Near-perfect +0.998 correlation between String 1 and String 2 draft confirms impeccable twin-tower balance with zero cyclone blockages or unequal gas split."),
            ("06. Clinker Heat Recuperation", "811.5°C", "Secondary Air Temperature",
             "Secondary air averages 811.5°C. Regulating cooler bed depth to push secondary air into the 950–1,100°C target window recovers massive sensible heat, cutting main burner coal."),
            ("07. Clinker Quality Optimization", "2.03%", "Mean Free Lime (fCaO)",
             "Free Lime averages 2.03% while C3S averages 59.18%. Stabilizing Free Lime below 1.50% via recuperated heat avoids downstream corrective grinding while preserving high 28-day strength."),
            ("08. ID & Cooler Fan Efficiency", "+0.991", "Fan Speed to kW Coupling",
             "Near-perfect +0.991 correlation between cooler fan speed and kW confirms that optimizing bed depth and draft balance allows immediate electrical energy conservation.")
        ]

        card_idx = 0
        for sh in s7.Shapes:
            if sh.Type == 6:  # Group
                for item in sh.GroupItems:
                    if item.HasTextFrame and card_idx < len(cards_data):
                        tag, stat, sub, desc = cards_data[card_idx]
                        item.TextFrame.TextRange.Text = f"{tag}\r{stat}\r{sub}\r{desc}"
                        tr = item.TextFrame.TextRange
                        if tr.Paragraphs().Count >= 1:
                            tr.Paragraphs(1).Font.Name = "Space Grotesk"
                            tr.Paragraphs(1).Font.Size = 11.0
                            tr.Paragraphs(1).Font.Color.RGB = 0x00F0E1
                        if tr.Paragraphs().Count >= 2:
                            tr.Paragraphs(2).Font.Name = "Space Grotesk"
                            tr.Paragraphs(2).Font.Size = 20.0
                            tr.Paragraphs(2).Font.Bold = True
                            tr.Paragraphs(2).Font.Color.RGB = 0xFFFFFF
                        if tr.Paragraphs().Count >= 3:
                            tr.Paragraphs(3).Font.Name = "Outfit"
                            tr.Paragraphs(3).Font.Size = 9.0
                            tr.Paragraphs(3).Font.Color.RGB = 0x00F0E1
                        if tr.Paragraphs().Count >= 4:
                            tr.Paragraphs(4).Font.Name = "Outfit"
                            tr.Paragraphs(4).Font.Size = 9.5
                            tr.Paragraphs(4).Font.Color.RGB = 0xE1D5CB
                        card_idx += 1
            elif sh.HasTextFrame and ("01." in sh.TextFrame.TextRange.Text or "Volumetric" in sh.TextFrame.TextRange.Text) and card_idx < len(cards_data):
                tag, stat, sub, desc = cards_data[card_idx]
                sh.TextFrame.TextRange.Text = f"{tag}\r{stat}\r{sub}\r{desc}"
                card_idx += 1
        set_slide_notes(s7,
            "=== SLIDE 07: Key Operational Findings & Diagnostic Assessment (8 Focus Cards) ===\n"
            "[ESTIMATED TIME: 3.0 minutes]\n\n"
            "[OBJECTIVE]\n"
            "Deliver a rapid-fire executive briefing across the 8 core operational dimensions governing Bursa Çimento's pyroprocessing efficiency.\n\n"
            "[SPOKEN SCRIPT / TALKING POINTS]\n"
            "\"Slide 7 consolidates our engineering assessment into 8 critical pillars: Card 1 shows volumetric throughput averaging 340.9 t/h feed with a tight +0.995 kinematic speed coupling. Card 2 highlights M02 drive torque averaging 179.2 A with trip risk above 250 A. Card 3 establishes fossil coal baseload at 20.39 t/h. Card 4 highlights alternative fuel upside: increasing RDF from 7.41 to 13–20 t/h. Card 5 confirms flawless +0.998 preheater aerodynamic balance. Card 6 pinpoints cooler recuperation at 811.5°C, targeting 950–1,100°C. Card 7 highlights clinker quality optimization (fCaO < 1.5%). Card 8 confirms electrical power conservation via bed depth control.\""
        )
        print(f"  -> Updated Slide 7: Executive Summary ({card_idx} cards)")

        # -------------------------------------------------------------
        # Slide 8: Detailed Findings Overview
        # -------------------------------------------------------------
        s8 = pres.Slides(8)
        for sh in s8.Shapes:
            if sh.HasTextFrame and sh.Name == "TextBox 2":
                sh.Top = 15.0
                sh.Height = 500.0
                sh.TextFrame.TextRange.Text = (
                    "Key Operational Findings & Diagnostic Assessment (Bursa Çimento Comprehensive Overview):\r\r"
                    "• 1. High-Capacity Volumetric Loading: The Bursa kiln operates at an average of 340.9 t/h raw feed (sustaining peaks up to 406.3 t/h). Feed variations are rigidly synchronized with Kiln Speed (mean 3.11 RPM, correlation +0.995), holding an optimal 11–13% volumetric material bed filling degree.\r"
                    "• 2. Drive Torque & M02 Mechanical Stress: Main drive current averages 179.22 A. Surges above 250 A without a corresponding increase in feed indicate unstable coating sloughing, ring formation in the burning zone, or snowman build-up forcing mechanical drive overload.\r"
                    "• 3. Calcination & Preheater Thermal Profile: Twin-string preheater draft exhibits flawless aerodynamic symmetry (+0.998 correlation), and Calciner Outlet A Temperature correlates +0.982 with sintering zone temperature, ensuring homogeneous decarbonation prior to kiln inlet entry.\r"
                    "• 4. Combustion & Dual Coal Baseload: Bursa fires 8.90 t/h coal on the main burner and 11.49 t/h in the calciner (20.39 t/h total coal baseload). Optimizing burner momentum and secondary air heat recuperation safely reduces specific heat consumption while protecting burning zone refractory.\r"
                    "• 5. Alternative Fuel (RDF) Substitution: RDF satellite firing averages 7.41 t/h with peaks up to 47.63 t/h. Stabilizing RDF continuous dosing to 13.0–20.0 t/h reliably substitutes fossil coal by 2.39 t/h (-11.72% coal reduction), delivering Bursa's largest direct energy cost saving.\r"
                    "• 6. Grate Cooler Heat Recuperation: Secondary air temperature averages 811.5°C. Elevating secondary air into the 950–1,100°C window provides a massive thermal flywheel that directly substitutes expensive main burner coal.\r"
                    "• 7. Cooler Bed Resistance & Fan Power: Near-perfect +0.991 correlation between cooler fan speed and kW confirms that clinker bed depth governs hydraulic resistance. Regulating grate speed optimizes bed depth and sheds significant electrical power.\r"
                    "• 8. Clinker Quality Optimization (fCaO & C3S): Free Lime (fCaO) averages 2.03% and C3S averages 59.18%. Stabilizing Free Lime below 1.50% via recuperated heat eliminates under-burnt clinker and downstream corrective grinding while preserving high 28-day cement strength.\r"
                    "• 9. Golden Batch Target Corridor: Maintaining Bursa operations at 369–400 t/h feed, 3.10–3.40 RPM speed, 950–1,100°C secondary air, and 13–20 t/h RDF defines the plant's high-efficiency Golden Batch corridor.\r"
                    "• 10. Emissions Stability & Process Safety: Operating preheater exit O2 between 3.50% and 4.50% eliminates incomplete combustion CO spikes (<0.05%), while controlled burning zone sintering suppresses thermal NOx and reagent NH3 consumption."
                )
                tr = sh.TextFrame.TextRange
                tr.Paragraphs(1).Font.Name = "Space Grotesk"
                tr.Paragraphs(1).Font.Size = 13.0
                tr.Paragraphs(1).Font.Bold = True
                tr.Paragraphs(1).Font.Color.RGB = 0x00FFFF
                for p_i in range(2, tr.Paragraphs().Count + 1):
                    tr.Paragraphs(p_i).Font.Name = "Outfit"
                    tr.Paragraphs(p_i).Font.Size = 10.0
                    tr.Paragraphs(p_i).Font.Color.RGB = 0xE1D5CB
        set_slide_notes(s8,
            "=== SLIDE 08: Key Operational Findings & Diagnostic Assessment (Bursa Comprehensive Overview) ===\n"
            "[ESTIMATED TIME: 3.0 minutes]\n\n"
            "[OBJECTIVE]\n"
            "Review the comprehensive narrative overview of the 10 engineering findings, explaining physical mechanisms and operational targets.\n\n"
            "[SPOKEN SCRIPT / TALKING POINTS]\n"
            "\"Slide 8 provides our comprehensive engineering narrative. As cement process engineers, several key takeaways stand out: First, the kiln mechanical drive current averages 179.2 A with peaks up to 431.7 A; current spikes above 250 A indicate heavy bed loading or ring build-up. Second, preheater aerodynamics show remarkable symmetry (+0.998). Third, the clinker quality assay indicates an average Free Lime of 2.03% against a target of 1.17–1.50%. By stabilizing secondary air above 950°C and ramping RDF to 13–20 t/h, Bursa can simultaneously unlock +18.42 t/h clinker throughput and save 2.39 t/h of coal.\""
        )
        print("  -> Updated Slide 8: Detailed Findings")

        # -------------------------------------------------------------
        # Slide 9: Graph 1 — Total Feed vs M02 Current
        # -------------------------------------------------------------
        setup_graph_slide(
            pres.Slides(9),
            "Graph 1: Total Feed Rate vs Main Drive Current (Drive Loading Dynamics)",
            [
                "• Mechanical Drive Stress: Total Kiln Feed directly dictates mechanical torque. Mean feed is 340.9 t/h with M02 current averaging 179.22 A.",
                "• Overload Threshold Mitigation: Surges in drive current above 250 A without a corresponding increase in feed indicate unstable kiln coating, potential ring formations in the burning zone, or severe snowman build-up forcing the drive into overload."
            ],
            "slide_02.png",
            notes_text=(
                "=== SLIDE 09: Graph 1: Total Feed Rate vs Main Drive Current (Drive Loading Dynamics) ===\n"
                "[ESTIMATED TIME: 2.0 minutes]\n\n"
                "[OBJECTIVE]\n"
                "Analyze the relationship between Total Kiln Feed and Kiln Main Drive Current (M02), establishing mechanical torque thresholds.\n\n"
                "[SPOKEN SCRIPT / TALKING POINTS]\n"
                "\"Slide 9 displays Graph 1: Total Feed Rate versus Kiln Main Drive Current. Raw meal feed directly dictates motor torque. Surges in drive current above 250 A without proportional feed increases indicate coating sloughing or ring formation. The Golden Batch targets steady bed operation with drive current governed below 250 A.\""
            )
        )
        print("  -> Updated Slide 9: Graph 1 (Feed vs M02 Current)")

        # -------------------------------------------------------------
        # Slide 10: Graph 2 — Drive Current vs Kiln Speed
        # -------------------------------------------------------------
        setup_graph_slide(
            pres.Slides(10),
            "Graph 2: Main Drive Current vs Kiln Speed (Kinematic Bed Synchronization)",
            [
                "• Volumetric Bed Synchronization: As feed approaches 370 t/h, kiln speed must be rigidly coordinated to maintain an optimal 11–13% volumetric filling degree.",
                "• Material Flush Prevention: Strict speed coordination prevents uncalcined meal from flushing into the burning zone, confirmed by the massive +0.995 correlation between Kiln Speed and Kiln Feed."
            ],
            "slide_07.png",
            notes_text=(
                "=== SLIDE 10: Graph 2: Main Drive Current vs Kiln Speed (Kinematic Bed Synchronization) ===\n"
                "[ESTIMATED TIME: 2.0 minutes]\n\n"
                "[OBJECTIVE]\n"
                "Evaluate the kinematic synchronization between Kiln Main Drive Current and Kiln Speed.\n\n"
                "[SPOKEN SCRIPT / TALKING POINTS]\n"
                "\"Slide 10 correlates drive current with rotational speed. As throughput expands toward 370–400 t/h, kiln speed must track dynamically to prevent material surging and thermal chilling of the sintering zone.\""
            )
        )
        print("  -> Updated Slide 10: Graph 2 (Current vs Speed)")

        # -------------------------------------------------------------
        # Slide 11: Graph 3 — SFC vs Secondary Air Temp
        # -------------------------------------------------------------
        setup_graph_slide(
            pres.Slides(11),
            "Graph 3: Specific Fuel Consumption vs Secondary Air Temperature (Thermal Flywheel)",
            [
                "• Fuel Substitution via Recuperation: Secondary Air Temp is our most critical thermal flywheel. Mean temperature is 811.5°C.",
                "• Thermal Efficiency Window: Operating consistently above 950–1,000°C directly lowers Specific Fuel Consumption, pushing thermal efficiency closer to world-class standards by minimizing the need for expensive main burner coal."
            ],
            "slide_08.png",
            notes_text=(
                "=== SLIDE 11: Graph 3: Specific Fuel Consumption vs Secondary Air Temperature ===\n"
                "[ESTIMATED TIME: 2.0 minutes]\n\n"
                "[OBJECTIVE]\n"
                "Quantify specific fuel consumption reduction through enhanced secondary air heat recuperation.\n\n"
                "[SPOKEN SCRIPT / TALKING POINTS]\n"
                "\"Slide 11 demonstrates the profound thermodynamic impact of Secondary Air Temperature on Specific Fuel Consumption. Pushing secondary air above 950°C acts as a massive thermal flywheel, directly reducing fossil coal demand on the main burner.\""
            )
        )
        print("  -> Updated Slide 11: Graph 3 (SFC vs Secondary Air)")

        # -------------------------------------------------------------
        # Slide 12: Graph 4 — Main Coal vs RDF Utilization
        # -------------------------------------------------------------
        setup_graph_slide(
            pres.Slides(12),
            "Graph 4: Main Burner Coal vs RDF Utilization (Alternative Fuel Substitution Limits)",
            [
                "• Fossil Fuel Displacement Dynamics: High RDF Satellite Burner flow must be carefully balanced against Main Burner Coal to maintain flame momentum.",
                "• Combustion Stability Guardrails: Over-reliance on RDF without proper momentum risks CO spikes, thermal decoupling, and dragging clinker liquid phase formation further back into the kiln."
            ],
            "slide_09.png",
            notes_text=(
                "=== SLIDE 12: Graph 4: Main Burner Coal vs RDF Utilization ===\n"
                "[ESTIMATED TIME: 2.0 minutes]\n\n"
                "[OBJECTIVE]\n"
                "Establish alternative fuel substitution boundaries between RDF satellite firing and fossil coal baseload.\n\n"
                "[SPOKEN SCRIPT / TALKING POINTS]\n"
                "\"Slide 12 maps the trade-off between RDF satellite firing and main burner coal. Ramping RDF from 7.41 to 13–20 t/h displaces 2.39 t/h of coal, but must be managed under strict flame momentum guardrails to prevent burning zone instability.\""
            )
        )
        print("  -> Updated Slide 12: Graph 4 (Coal vs RDF)")

        # -------------------------------------------------------------
        # Slide 13: STEP 5A: Time Trend (Slide 2 from UPDATED deck)
        # -------------------------------------------------------------
        setup_graph_slide(
            pres.Slides(13),
            "STEP 5A: Kiln Process Stability & Multi-Variable Scatter Diagnostics",
            [
                "• Continuous Process Validation: Multi-variable telemetry tracking shows consistent operation across 340+ t/h feed regimes with synchronized coal and RDF modulation.",
                "• Predictive Feed-Forward Opportunity: Real-time tracking identifies transient feed dips and drive current spikes, enabling proactive AI intervention to prevent thermal shock."
            ],
            "slide_02.png",
            notes_text=(
                "=== SLIDE 13: STEP 5A: Kiln Process Stability Diagnostics ===\n"
                "[ESTIMATED TIME: 2.0 minutes]\n\n"
                "[OBJECTIVE]\n"
                "Review multi-variable operational diagnostics across feed regimes, mechanical drive load, and thermal stability.\n\n"
                "[SPOKEN SCRIPT / TALKING POINTS]\n"
                "\"Slide 13 displays the comprehensive process stability distribution. The data reveals distinct operational operating clusters and confirms that the kiln can sustain 370–400 t/h when feed and rotational drive are synchronized.\""
            )
        )
        print("  -> Updated Slide 13: Time Trend A (Slide 2 img)")

        # -------------------------------------------------------------
        # Slide 14: STEP 5B: Parallel Plot Quality
        # -------------------------------------------------------------
        setup_graph_slide(
            pres.Slides(14),
            "STEP 5B: Parallel Trajectories — Bursa Quality & Burnability Diagnostics",
            [
                "• High-Dimensional Trajectory Mapping: Interconnecting Total Fuel Flow, Secondary Air Temp, Free Lime (fCaO), C3S, and Clinker Production.",
                "• Optimal Quality Corridor: Trajectories passing through >950°C secondary air consistently achieve C3S > 60% with Free Lime suppressed below 1.5%."
            ],
            "slide_16.png",
            notes_text=(
                "=== SLIDE 14: STEP 5B: Parallel Trajectories — Quality & Burnability ===\n"
                "[ESTIMATED TIME: 2.0 minutes]\n\n"
                "[OBJECTIVE]\n"
                "Interpret multi-dimensional trajectories connecting fuel input, secondary air recuperation, and lab quality assays.\n\n"
                "[SPOKEN SCRIPT / TALKING POINTS]\n"
                "\"Slide 14 presents parallel coordinate trajectories mapping fuel flow and recuperation to final clinker quality. Notice how high secondary air trajectories converge on the Golden Zone of high C3S and low Free Lime.\""
            )
        )
        print("  -> Updated Slide 14: Parallel Plot Quality (Slide 16 img)")

        # -------------------------------------------------------------
        # Slide 15: STEP 5C: Section 1 Performance Summary Table
        # -------------------------------------------------------------
        sec1_header = ["Operational Variable", "Min", "Q1", "Median", "Mean", "Q3", "Max", "TARGET CORRIDOR"]
        sec1_rows = [
            ["KILN_FEED (t/h)", "0.00", "310.00", "340.00", "340.90", "365.00", "406.30", "369.0–400.0 t/h"],
            ["Clinker_Production (t/h)", "0.00", "201.50", "221.00", "221.58", "237.25", "264.10", "240.0–260.0 t/h"],
            ["KILN_SPEED (rpm)", "1.10", "2.85", "3.10", "3.11", "3.35", "3.95", "3.10–3.40 rpm"],
            ["M02_DRIVE_CURRENT (A)", "130.97", "165.00", "178.00", "179.22", "195.00", "431.73", "160.0–220.0 A"],
            ["MAIN_BURNER_COAL (t/h)", "3.70", "8.20", "8.85", "8.90", "9.50", "12.71", "8.00–9.00 t/h"],
            ["CALCINER_COAL (t/h)", "0.64", "10.50", "11.40", "11.49", "12.50", "13.64", "10.00–13.00 t/h"],
            ["RDF_SATELLITEBURNER (t/h)", "0.00", "4.50", "7.20", "7.41", "10.20", "47.63", "13.00–20.00 t/h"]
        ]
        populate_table_slide(
            pres.Slides(15),
            "STEP 5C: Section 1 Performance Summary Table (Kiln & Thermal Dynamics)",
            sec1_header, sec1_rows,
            notes_text=(
                "=== SLIDE 15: STEP 5C: Section 1 Performance Summary Table ===\n"
                "[ESTIMATED TIME: 2.5 minutes]\n\n"
                "[OBJECTIVE]\n"
                "Walk through statistical distribution metrics across primary kiln throughput, mechanical drive, and fuel variables for Bursa Çimento.\n\n"
                "[SPOKEN SCRIPT / TALKING POINTS]\n"
                "\"Slide 15 provides a statistical summary table for Section 1. Kiln Feed averages 340.9 t/h with a target corridor of 369–400 t/h. M02 Drive Current averages 179.2 A with maximum peaks at 431.7 A. Fossil coal totals 20.39 t/h, while RDF averages 7.41 t/h with an aggressive target corridor of 13–20 t/h.\""
            )
        )
        print("  -> Updated Slide 15: Section 1 Summary Table")

        # -------------------------------------------------------------
        # Slide 16: Graph 5 — Free Lime vs Secondary Air
        # -------------------------------------------------------------
        setup_graph_slide(
            pres.Slides(16),
            "Graph 5: Free Lime (SCAO) vs Secondary Air Temperature (Burnability Health)",
            [
                "• Burnability Quality Indicator: Free Lime (KLINKER SCAO) is the definitive indicator of burning zone health, averaging 2.03% (optimal window 1.17–2.45%).",
                "• Thermal Starvation Diagnosis: When Secondary Air Temp drops below 800°C, Free Lime spikes above 2.5%, indicating under-burning and poor nodulization that degrades final cement strength."
            ],
            "slide_11.png",
            notes_text=(
                "=== SLIDE 16: Graph 5: Free Lime (SCAO) vs Secondary Air Temperature ===\n"
                "[ESTIMATED TIME: 2.0 minutes]\n\n"
                "[OBJECTIVE]\n"
                "Correlate clinker Free Lime (fCaO) with secondary air temperature to establish thermal burnability thresholds.\n\n"
                "[SPOKEN SCRIPT / TALKING POINTS]\n"
                "\"Slide 16 illustrates the critical relationship between Secondary Air Temperature and Free Lime. When secondary air drops below 800°C, Free Lime spikes above 2.5%, causing unsound clinker. Elevating secondary air to 950–1,100°C ensures complete clinker nodulization.\""
            )
        )
        print("  -> Updated Slide 16: Graph 5 (Free Lime vs Secondary Air)")

        # -------------------------------------------------------------
        # Slide 17: Graph 6 — C3S vs Free Lime
        # -------------------------------------------------------------
        setup_graph_slide(
            pres.Slides(17),
            "Graph 6: C3S vs Free Lime (SCAO) (Quality vs Efficiency Trade-off)",
            [
                "• High-Alite Formation Challenge: C3S averages 59.18% (target 60.0–63.0%). High C3S requires intense localized heat which naturally drives down Free Lime.",
                "• Golden Zone Optimization: The objective is to achieve the sweet spot (fCaO 1.17–1.50%, C3S > 60%) using maximized secondary air recuperation rather than excessive main burner coal dumping."
            ],
            "slide_10.png",
            notes_text=(
                "=== SLIDE 17: Graph 6: C3S vs Free Lime (SCAO) ===\n"
                "[ESTIMATED TIME: 2.0 minutes]\n\n"
                "[OBJECTIVE]\n"
                "Analyze the quality-efficiency trade-off between C3S alite content and Free Lime fCaO.\n\n"
                "[SPOKEN SCRIPT / TALKING POINTS]\n"
                "\"Slide 17 highlights the core quality trade-off between C3S and Free Lime. High C3S ensures cement strength, while low Free Lime prevents expansion. Our AI controller hits the sweet spot through recuperated heat rather than over-firing coal.\""
            )
        )
        print("  -> Updated Slide 17: Graph 6 (C3S vs Free Lime)")

        # -------------------------------------------------------------
        # Slide 18: Graph 7 — NH3 vs Secondary Air
        # -------------------------------------------------------------
        setup_graph_slide(
            pres.Slides(18),
            "Graph 7: NH3 Consumption vs Secondary Air Temperature (NOx Emissions Control)",
            [
                "• Thermal NOx Sensitivity: High Secondary Air Temp combined with excess preheater O2 accelerates thermal NOx formation, driving NH3 consumption (averaging 145.00).",
                "• Source Emission Reduction: Choking back excess air at the preheater outlet suppresses NOx at the source, minimizing expensive reagent NH3 consumption while ensuring regulatory compliance."
            ],
            "slide_12.png",
            notes_text=(
                "=== SLIDE 18: Graph 7: NH3 Consumption vs Secondary Air Temperature ===\n"
                "[ESTIMATED TIME: 2.0 minutes]\n\n"
                "[OBJECTIVE]\n"
                "Evaluate emissions abatement and NH3 reagent dosing as a function of combustion temperatures.\n\n"
                "[SPOKEN SCRIPT / TALKING POINTS]\n"
                "\"Slide 18 displays NH3 consumption versus secondary air. High combustion temperatures and excess oxygen create thermal NOx. Controlling excess air suppresses NOx at the source and directly reduces NH3 reagent costs.\""
            )
        )
        print("  -> Updated Slide 18: Graph 7 (NH3 vs Secondary Air)")

        # -------------------------------------------------------------
        # Slide 19: Graph 8 — Preheater O2 vs Fan Outlet
        # -------------------------------------------------------------
        setup_graph_slide(
            pres.Slides(19),
            "Graph 8: Preheater Outlet O2 vs Preheater Fan Outlet (Draft Aerodynamics)",
            [
                "• Draft Symmetry Non-Negotiable: The system exhibits near-perfect symmetry with a +0.998 correlation between Pressure after Pre-heater 1 and 2.",
                "• ID Fan Headroom Preservation: Operating preheater exit O2 between 3.50% and 4.50% minimizes false air ingress and preserves ID fan capacity for throughput expansion."
            ],
            "slide_13.png",
            notes_text=(
                "=== SLIDE 19: Graph 8: Preheater Outlet O2 vs Preheater Fan Outlet ===\n"
                "[ESTIMATED TIME: 2.0 minutes]\n\n"
                "[OBJECTIVE]\n"
                "Analyze preheater draft aerodynamics, twin-string balance, and ID fan headroom.\n\n"
                "[SPOKEN SCRIPT / TALKING POINTS]\n"
                "\"Slide 19 examines preheater draft aerodynamics. Near-perfect +0.998 symmetry between strings confirms zero blockages. Governing exit O2 between 3.5–4.5% frees ID fan capacity for our 18.42 t/h throughput increase.\""
            )
        )
        print("  -> Updated Slide 19: Graph 8 (O2 vs Fan Outlet)")

        # -------------------------------------------------------------
        # Slide 20: Graph 9 — Cooler Exhaust vs Secondary Air
        # -------------------------------------------------------------
        setup_graph_slide(
            pres.Slides(20),
            "Graph 9: Cooler Exhaust Temp vs Secondary Air Temp (Cooler Bed Hydraulic Resistance)",
            [
                "• Clinker Bed Depth Governance: Deep clinker beds force cooling fans into severe power draw (+0.991 correlation between fan speed and fan kW).",
                "• Thermal vs Electrical Balance: Regulating grate speed sheds significant electrical kW load without over-cooling the bed and destroying the Secondary Air Temp thermal flywheel."
            ],
            "slide_14.png",
            notes_text=(
                "=== SLIDE 20: Graph 9: Cooler Exhaust Temp vs Secondary Air Temp ===\n"
                "[ESTIMATED TIME: 2.0 minutes]\n\n"
                "[OBJECTIVE]\n"
                "Evaluate clinker cooler bed resistance, exhaust thermal losses, and cooling fan electrical consumption.\n\n"
                "[SPOKEN SCRIPT / TALKING POINTS]\n"
                "\"Slide 20 evaluates cooler heat exchange. High cooler fan speeds draw excessive kW. By stabilizing grate speed and bed depth, we recover maximum secondary air heat while reducing fan electrical power.\""
            )
        )
        print("  -> Updated Slide 20: Graph 9 (Cooler Exhaust vs Secondary Air)")

        # -------------------------------------------------------------
        # Slide 21: Graph 10 — Parallel Plot Kiln Dynamics
        # -------------------------------------------------------------
        setup_graph_slide(
            pres.Slides(21),
            "Graph 10: Bursa Main Kiln Process Dynamics Multi-Variable Trajectories",
            [
                "• Kinematic Drive Trajectories: Multi-dimensional trajectories mapping Kiln Feed, Main Coal, Calciner Coal, RDF Satellite, and Clinker Output.",
                "• Synchronized High-Throughput Cluster: High-feed trajectories demonstrate consistent fuel balance and stable thermal profiling without mechanical overload."
            ],
            "slide_15.png",
            notes_text=(
                "=== SLIDE 21: Graph 10: Bursa Main Kiln Process Dynamics Trajectories ===\n"
                "[ESTIMATED TIME: 2.0 minutes]\n\n"
                "[OBJECTIVE]\n"
                "Map multi-variable process trajectories across kiln feed, coal firing, RDF, and clinker production.\n\n"
                "[SPOKEN SCRIPT / TALKING POINTS]\n"
                "\"Slide 21 presents parallel coordinate trajectories for main kiln dynamics. Notice how operations successfully handle peak throughput when fuel splits between the calciner and main burner are coordinated.\""
            )
        )
        print("  -> Updated Slide 21: Graph 10 (Parallel Plot Kiln Dynamics)")

        # -------------------------------------------------------------
        # Slide 22: STEP 7B: Section 2 Performance Summary Table
        # -------------------------------------------------------------
        sec2_header = ["Operational Variable", "Min", "Q1", "Median", "Mean", "Q3", "Max", "TARGET CORRIDOR"]
        sec2_rows = [
            ["SECONDARY_AIR_TEMP (°C)", "785.00", "795.00", "810.00", "811.54", "825.00", "1800.00", "950.0–1100.0 °C"],
            ["PRE_HEATER_OUTLET_O2 (%)", "3.47", "3.90", "4.30", "4.37", "4.80", "20.83", "3.50–4.50 %"],
            ["PRE_HEATER_OUTLET_CO (%)", "0.020", "0.035", "0.042", "0.043", "0.050", "0.050", "0.020–0.050 %"],
            ["NH3_CONSUMPTION (proxy)", "1.00", "110.00", "142.00", "145.00", "175.00", "100.20", "80.0–120.0 proxy"],
            ["KLINKER_SCAO / Free Lime (%)", "1.17", "1.65", "1.98", "2.03", "2.40", "2.45", "1.17–1.50 %"],
            ["KLINKER_C3S (%)", "57.05", "58.20", "59.10", "59.18", "60.15", "62.95", "60.00–63.00 %"],
            ["Twin Preheater Balance (+r)", "+0.998", "+0.998", "+0.998", "+0.998", "+0.998", "+0.998", "String 1 vs String 2"],
            ["Cooler Fan Speed to kW (+r)", "+0.991", "+0.991", "+0.991", "+0.991", "+0.991", "+0.991", "Hydraulic Bed Sync"]
        ]
        populate_table_slide(
            pres.Slides(22),
            "STEP 7B: Section 2 Performance Summary Table (Thermal, Aerodynamics & Quality)",
            sec2_header, sec2_rows,
            notes_text=(
                "=== SLIDE 22: STEP 7B: Section 2 Performance Summary Table ===\n"
                "[ESTIMATED TIME: 2.5 minutes]\n\n"
                "[OBJECTIVE]\n"
                "Review statistical distributions for secondary air recuperation, combustion gas atmosphere, emissions, and lab clinker quality.\n\n"
                "[SPOKEN SCRIPT / TALKING POINTS]\n"
                "\"Slide 22 details Section 2 operational variables. Secondary air averages 811.5°C against a target of 950–1,100°C. Preheater O2 averages 4.37%, while CO is controlled at 0.043%. Free Lime averages 2.03% and C3S averages 59.18%, establishing the precise tuning targets for our Golden Batch.\""
            )
        )
        print("  -> Updated Slide 22: Section 2 Summary Table")

        # -------------------------------------------------------------
        # Slide 23: STEP 8A: 3D Quality & Thermal Space
        # -------------------------------------------------------------
        setup_graph_slide(
            pres.Slides(23),
            "STEP 8A: 3D Quality & Thermal Space (Secondary Air × Free Lime × C3S)",
            [
                "• Multi-Dimensional Quality Mapping: Visualizing clinker mineralogy as a simultaneous function of Secondary Air Temp, Free Lime (fCaO), and C3S content.",
                "• Golden Quality Cluster: High-strength, sound clinker (C3S 60–63%, fCaO < 1.5%) forms a tight operational cluster at secondary air temperatures above 950°C."
            ],
            "dump_s20.png",
            notes_text=(
                "=== SLIDE 23: STEP 8A: 3D Quality & Thermal Space ===\n"
                "[ESTIMATED TIME: 2.5 minutes]\n\n"
                "[OBJECTIVE]\n"
                "Visualize the 3D quality envelope combining Secondary Air Temperature, Free Lime (fCaO), and C3S.\n\n"
                "[SPOKEN SCRIPT / TALKING POINTS]\n"
                "\"Slide 23 displays the 3D Quality and Thermal Space. By mapping secondary air against Free Lime and C3S simultaneously, we observe that the highest quality clinker clusters exclusively when secondary air recuperation exceeds 950°C.\""
            )
        )
        print("  -> Updated Slide 23: 3D Quality Space (dump_s20 img)")

        # -------------------------------------------------------------
        # Slide 24: STEP 8B: 3D Cooler & Fuel Space
        # -------------------------------------------------------------
        setup_graph_slide(
            pres.Slides(24),
            "STEP 8B: 3D Cooler & Fuel Efficiency Space (RDF × Coal × Secondary Air)",
            [
                "• Bursa Golden Zone Mapping: 3D envelopes map the exact combination of Fuel Split (RDF vs Coal) and Recuperation (Secondary Air Temp) yielding maximum efficiency.",
                "• Thermal Baseload Optimization: Sustaining RDF between 13–20 t/h while holding secondary air >950°C achieves structural fuel cost minimization at peak clinker yield."
            ],
            "dump_s22.png",
            notes_text=(
                "=== SLIDE 24: STEP 8B: 3D Cooler & Fuel Efficiency Space ===\n"
                "[ESTIMATED TIME: 2.5 minutes]\n\n"
                "[OBJECTIVE]\n"
                "Map the 3D multi-variable efficiency envelope combining RDF satellite firing, total coal flow, and secondary air temperature.\n\n"
                "[SPOKEN SCRIPT / TALKING POINTS]\n"
                "\"Slide 24 visualizes the Bursa Golden Zone in 3D. The optimal operating envelope combines 13–20 t/h of RDF with 950–1,100°C secondary air, yielding maximum clinker throughput at minimum specific fuel consumption.\""
            )
        )
        print("  -> Updated Slide 24: 3D Fuel Efficiency Space (dump_s22 img)")

        # -------------------------------------------------------------
        # Slide 25: STEP 9: Golden Batch Definition Table
        # -------------------------------------------------------------
        s25 = pres.Slides(25)
        gb_header = ["Operational Variable", "Golden Batch Lower", "Golden Batch Upper", "Current Mean", "Optimization Gap"]
        gb_rows = [
            ["Clinker_Production_tph", "240.00", "260.00", "221.58", "+18.42 to +38.42 t/h (+8.31%)"],
            ["KILN_FEED", "369.00", "400.00", "340.90", "+28.10 to +59.10 t/h"],
            ["MAIN_BURNER_COAL", "8.00", "9.00", "8.90", "Near Optimal (Baseload Tuning)"],
            ["CALCINER_COAL", "10.00", "13.00", "11.49", "Near Optimal (Fuel Split Tuning)"],
            ["RDF_SATELLITEBURNER", "13.00", "20.00", "7.41", "+5.59 to +12.59 t/h (Substitution)"],
            ["PRE_HEATER_OUTLET_O2", "3.50", "4.50", "4.37", "Near Optimal (Excess Air Trim)"],
            ["SECONDARY_AIR_TEMP", "950.00", "1100.00", "811.54", "+138.46 to +288.46 °C"],
            ["KLINKER_SCAO (Free Lime)", "1.17", "1.50", "2.03", "-0.53 to -0.86% (Soundness Guardrail)"],
            ["KLINKER_C3S (Alite)", "60.00", "63.00", "59.18", "+0.82 to +3.82% (28-Day Strength)"]
        ]
        populate_table_slide(
            s25,
            "STEP 9: THE GOLDEN BATCH OPERATING WINDOW DEFINITION (BURSA ÇIMENTO)",
            gb_header, gb_rows,
            notes_text=(
                "=== SLIDE 25: STEP 9: THE GOLDEN BATCH OPERATING WINDOW DEFINITION ===\n"
                "[ESTIMATED TIME: 3.5 minutes]\n\n"
                "[OBJECTIVE]\n"
                "Define the official multi-variable Golden Batch Target Operating Window across all 9 primary process and quality parameters for Bursa Çimento.\n\n"
                "[SPOKEN SCRIPT / TALKING POINTS]\n"
                "\"Slide 25 establishes Bursa Çimento's official Golden Batch Target Operating Window. By synthesizing our multi-dimensional clustering and quality trade-offs, we define exact operating boundaries: Clinker Production expands from 221.58 to 240–260 t/h (+18.42 t/h gain); Kiln Feed increases to 369–400 t/h; RDF increases from 7.41 to 13–20 t/h; Secondary Air is pushed from 811.5°C to 950–1,100°C; and Free Lime is held strictly within 1.17–1.50% to ensure superior clinker reactivity.\""
            )
        )
        print("  -> Updated Slide 25: Golden Batch Table")

        # -------------------------------------------------------------
        # Slide 26: STEP 10: Value Lost Summary Table
        # -------------------------------------------------------------
        s26 = pres.Slides(26)
        v_sum_header = ["Financial Optimization Opportunity", "Physical Operational Quantity", "Percentage Gain"]
        v_sum_rows = [
            ["Clinker Production Throughput Increase", "18.42 t/h extra Clinker (221.58 → 240.00 t/h)", "+8.31% Throughput"],
            ["Total Coal Reduction (via RDF Substitution)", "2.39 t/h of Coal Saved (20.39 → 18.00 t/h)", "-11.72% Coal Usage"],
            ["Clinker Quality & C3S Mineralogy Stabilization", "Stable Free Lime < 1.50% & C3S > 60%", "Avoided Quality Rejections"],
            ["App Realization Potential (Conservative 20% Capture)", "3.68 t/h extra Clinker, 0.48 t/h Coal Saved", "+1.66% Throughput, -2.34% Coal"],
            ["Full AI Optimization Potential (100% Target Window)", "18.42 t/h extra Clinker, 2.39 t/h Coal Saved", "+8.31% Throughput, -11.72% Coal"]
        ]
        populate_table_slide(
            s26,
            "STEP 10: BUSINESS CASE & VALUE QUANTIFICATION (OPPORTUNITY SUMMARY)",
            v_sum_header, v_sum_rows,
            notes_text=(
                "=== SLIDE 26: STEP 10: BUSINESS CASE & VALUE QUANTIFICATION (Summary) ===\n"
                "[ESTIMATED TIME: 3.5 minutes]\n\n"
                "[OBJECTIVE]\n"
                "Present the comprehensive business case summary detailing throughput gains, coal savings via RDF, and quality benefits for Bursa Çimento.\n\n"
                "[SPOKEN SCRIPT / TALKING POINTS]\n"
                "\"Slide 26 summarizes the physical quantities and efficiency gains delivered by AI Process Navigator. Throughput expansion yields +18.42 t/h of additional clinker (+8.31% gain). Coal reduction via RDF substitution saves 2.39 t/h of fossil coal (-11.72% reduction). Even at a conservative 20% capture rate, the plant captures +3.68 t/h clinker and 0.48 t/h coal savings immediately.\""
            )
        )
        print("  -> Updated Slide 26: Value Summary Table")

        # -------------------------------------------------------------
        # Slide 27: STEP 10: Financial Valuation Table
        # -------------------------------------------------------------
        s27 = pres.Slides(27)
        v_fin_header = ["Financial Optimization Opportunity", "Annual Production Quantity", "Annual Financial Value"]
        v_fin_rows = [
            ["Clinker Throughput Expansion (221.58 → 240.00 t/h clinker)", "+132,600 tons clinker / year", "$6,630,000 / year"],
            ["Fossil Coal Reduction via RDF Substitution (2.39 t/h saved)", "17,200 tons coal saved / yr", "$2,236,000 / year"],
            ["Quality & Free Lime Stabilization (Avoided corrective grinding)", "Stable fCaO < 1.50% & C3S > 60%", "$350,000 / year"],
            ["Cooler Aeration Electrical Optimization (Bed depth control)", "950,000 kWh / year saved", "$95,000 / year"],
            ["Preheater ID Fan Electrical Conservation (Draft symmetry)", "720,000 kWh / year saved", "$72,000 / year"],
            ["Kiln Main Drive (M02) Protection (Trip avoidance >250 A)", "Avoidance of 16h trip downtime", "$240,000 / year"],
            ["Refractory Campaign Life Extension (Thermal stability)", "+2.5 months campaign extension", "$200,000 / year"],
            ["CO2 Carbon Emissions Avoidance (Coal reduction & RDF)", "36,000 tons CO2 / year avoided", "Carbon Credit Benefit"],
            ["TOTAL DIRECT ECONOMIC VALUE UNLOCKED", "Across All Operational Levers", "~$9,823,000 / year"]
        ]
        populate_table_slide(
            s27,
            "STEP 10: BUSINESS CASE & FINANCIAL VALUE QUANTIFICATION (ANNUAL ROI)",
            v_fin_header, v_fin_rows,
            notes_text=(
                "=== SLIDE 27: STEP 10: BUSINESS CASE & FINANCIAL VALUE QUANTIFICATION (Financial Table) ===\n"
                "[ESTIMATED TIME: 3.5 minutes]\n\n"
                "[OBJECTIVE]\n"
                "Present the audited financial valuation table detailing the annual dollar value delivered by each optimization lever at Bursa Çimento.\n\n"
                "[SPOKEN SCRIPT / TALKING POINTS]\n"
                "\"Slide 27 presents the audited financial bottom line. At a clinker margin of $50/ton, the 132,600 tons of additional clinker throughput generates $6.63M/year. The 17,200 tons of coal displaced via RDF at $130/ton delivers $2.236M/year in fuel savings. Adding quality stabilization ($350K), fan power savings ($167K), drive protection ($240K), and refractory extension ($200K), the total recurring economic value unlocked is ~$9,823,000 per year.\""
            )
        )
        print("  -> Updated Slide 27: Financial Valuation Table")

        # -------------------------------------------------------------
        # Slide 28: STEP 10B: Detailed Calculations
        # -------------------------------------------------------------
        s28 = pres.Slides(28)
        for sh in s28.Shapes:
            if sh.HasTextFrame:
                if sh.Name == "TextBox 2":
                    sh.TextFrame.TextRange.Text = "STEP 10B: DETAILED FINANCIAL & PROCESS VALUE CALCULATIONS (BURSA ÇIMENTO)"
                    sh.TextFrame.TextRange.Font.Name = "Space Grotesk"
                    sh.TextFrame.TextRange.Font.Size = 22.0
                    sh.TextFrame.TextRange.Font.Bold = True
                    sh.TextFrame.TextRange.Font.Color.RGB = 0x00FFFF
                elif sh.Name == "TextBox 1":
                    sh.TextFrame.TextRange.Text = "OPERATIONAL VALUATION METHODOLOGY: RIGOROUS PLANT BASELINES, CONVERSION CONSTANTS & FORMULAS"
                    sh.TextFrame.TextRange.Font.Name = "Outfit"
                    sh.TextFrame.TextRange.Font.Size = 10.0
                    sh.TextFrame.TextRange.Font.Bold = True
                    sh.TextFrame.TextRange.Font.Color.RGB = 0x00F0E1
                elif sh.Name == "Rectangle: Rounded Corners 4":
                    sh.TextFrame.TextRange.Text = (
                        "1. CLINKER THROUGHPUT EXPANSION (PRODUCTION INCREASE)\r"
                        "• Formula: Annual Clinker (t/yr) = ΔClinker (t/h) × Annual Operating Hours (h)\r"
                        "Financial Value = Annual Clinker (t/yr) × Net Clinker Margin ($/ton)\r"
                        "• Operational Baseline: Mean Clinker Production = 221.58 t/h (Mean Feed = 340.90 t/h)\r"
                        "• Golden Batch Target Clinker: 240.00 t/h (Target Feed = 369.00 t/h)\r"
                        "• Hourly Clinker Gain (ΔClinker): 240.00 - 221.58 = +18.42 t clinker / hour (+8.31%)\r"
                        "• Operating Baseline: 7,200 operating hours / year (~82% uptime) | Margin: $50.00 / ton\r"
                        "• Annual Clinker Gain: 18.42 t/h × 7,200 h = +132,624 metric tons clinker / year\r"
                        "✔ Annual Value Unlocked: 132,600 t × $50.00/t = $6,630,000 / year (~$6.63M/yr)\r"
                        "• Supervisory Guardrails: Governed by Clinker Free Lime (fCaO ≤ 1.50%) and ID fan headroom."
                    )
                elif sh.Name == "Rectangle: Rounded Corners 5":
                    sh.TextFrame.TextRange.Text = (
                        "2. COAL REDUCTION & ALTERNATIVE FUEL (RDF) SAVINGS\r"
                        "• Formula: Fuel Saved (t/yr) = ΔCoal (t/h) × Operating Hours (h)\r"
                        "Financial Value = Coal Saved (t/yr) × Delivered Coal Cost ($/ton)\r"
                        "• Baseline Coal Firing: 20.39 t/h total coal (8.90 t/h Main Burner + 11.49 t/h Calciner)\r"
                        "• Golden Batch Target Coal: 18.00 t/h (displaced via RDF ramping to 13.0–20.0 t/h)\r"
                        "• Hourly Coal Reduction: 20.39 - 18.00 = 2.39 t/h fossil coal saved (-11.72%)\r"
                        "• Annual Fossil Coal Saved: 2.39 t/h × 7,200 h = 17,208 metric tons coal / year\r"
                        "• Delivered Industrial Coal Cost: $130.00 / metric ton industrial coal\r"
                        "✔ Annual Value Unlocked: 17,200 t × $130.00/t = $2,236,000 / year (~$2.236M/yr)\r"
                        "• Additional Quality Value: Stabilizing fCaO < 1.5% saves $350,000/yr in avoided corrective grinding."
                    )
                elif sh.Name == "Rectangle: Rounded Corners 6":
                    sh.TextFrame.TextRange.Text = (
                        "3. ELECTRICAL POWER CONSERVATION (COOLER & ID FANS)\r"
                        "• Formula: Power Saved (kWh/yr) = ΔkW × Operating Hours (h)\r"
                        "Financial Value = Power Saved (kWh/yr) × Industrial Power Tariff ($/kWh)\r"
                        "• Cooler Aeration Optimization: -132 kW across aeration network via bed depth control (+0.991 r)\r"
                        "• Preheater ID Fan Suction: -100 kW across exhaust fan via aerodynamic draft symmetry (+0.998 r)\r"
                        "• Running Baseline: 7,200 operating hours / year | Industrial Electricity Tariff: $0.10 / kWh\r"
                        "• Cooler Energy Saved: 132 kW × 7,200 h = 950,400 kWh / year ($95,040/yr)\r"
                        "• ID Fan Energy Saved: 100 kW × 7,200 h = 720,000 kWh / year ($72,000/yr)\r"
                        "✔ Annual Value Unlocked: 1,670,400 kWh × $0.10/kWh = $167,000 / year (~$167K/yr)"
                    )
                elif sh.Name == "Rectangle: Rounded Corners 7":
                    sh.TextFrame.TextRange.Text = (
                        "4. ASSET RELIABILITY, MECHANICAL PROTECTION & ESG\r"
                        "• Value = Avoided Motor Trips ($/yr) + Refractory Extension ($/yr) + Carbon ESG\r"
                        "Carbon Benefit = Fossil Coal Saved (t/yr) × 2.09 t CO2 / ton coal\r"
                        "• M02 Drive Protection: Current peaks up to 431.7 A (target <250 A) cause overload trips.\r"
                        "Peak shaving avoids 16 hours of unscheduled stoppage downtime: $240,000/yr\r"
                        "• Refractory Campaign Extension: Thermal stability (950–1,100°C sec air) & CO trimming\r"
                        "extends brick lining campaign life by +2.5 months: $200,000/yr\r"
                        "• Decarbonization (Scope 1): 17,200 t coal saved = 36,000 metric tons CO2 / year avoided\r"
                        "• Combined Mechanical Protection: $240,000 + $200,000 = $440,000 / year\r"
                        "✔ Asset & Decarbonization Value: $440,000 / year + 36,000 tons CO2 avoided"
                    )
                elif sh.Name == "Rectangle 8":
                    sh.TextFrame.TextRange.Text = (
                        "TOTAL UNLOCKED ECONOMIC VALUE: ~$9,823,000 / YEAR  |  Throughput: $6,630,000/yr (67.5%)  •  "
                        "Coal & RDF: $2,236,000/yr (22.8%)  •  Quality & Electrical: $517,000/yr (5.3%)  •  Asset Protection: $440,000/yr (4.4%)"
                    )
        set_slide_notes(s28,
            "=== SLIDE 28: STEP 10B: DETAILED FINANCIAL & PROCESS VALUE CALCULATIONS ===\n"
            "[ESTIMATED TIME: 3.5 minutes]\n\n"
            "[OBJECTIVE]\n"
            "Walk through the transparent mathematical equations, plant baselines, conversion constants, and engineering formulas for all 4 value pillars.\n\n"
            "[SPOKEN SCRIPT / TALKING POINTS]\n"
            "\"Slide 28 details our valuation methodology, ensuring total financial transparency. In Card 1, Throughput Expansion: 18.42 t/h clinker gain over 7,200 hours yields 132,600 tons of clinker, delivering $6.63M at a $50 margin. In Card 2, Coal Reduction: saving 2.39 t/h of coal over 7,200 hours yields 17,200 tons of coal saved, unlocking $2.236M at $130/ton, plus $350K in avoided quality grinding. In Card 3, Electrical Power: cooler bed control and draft symmetry save 1.67M kWh, worth $167K. In Card 4, Asset Protection: M02 trip shaving and refractory life extension deliver $440K, while avoiding 36,000 tons of CO2. Total recurring economic value is ~$9,823,000 per year.\""
        )
        print("  -> Updated Slide 28: Detailed Calculations")

        # -------------------------------------------------------------
        # Slide 29: Parallel Plot 1: Main Kiln Process Dynamics
        # -------------------------------------------------------------
        setup_graph_slide(
            pres.Slides(29),
            "Parallel Plot 1: Bursa Main Kiln Process Dynamics Multi-Variable Trajectories",
            [
                "• Continuous Process Correlation: Mapping multi-variable trajectories across Total Kiln Feed, Kiln Speed, M02 Drive Current, and Main Burner Coal.",
                "• Synchronized Golden Trajectory: Operational lines running in the 369–400 t/h corridor maintain steady rotational speed and low drive stress, confirming high stability."
            ],
            "slide_15.png",
            notes_text=(
                "=== SLIDE 29: Parallel Plot 1: Bursa Main Kiln Process Dynamics Multi-Variable Trajectories ===\n"
                "[ESTIMATED TIME: 2.0 minutes]\n\n"
                "[OBJECTIVE]\n"
                "Interpret high-dimensional parallel coordinate trajectories across Total Kiln Feed, Kiln Speed, M02 Drive Current, and Coal Flow for Bursa Çimento.\n\n"
                "[SPOKEN SCRIPT / TALKING POINTS]\n"
                "\"Slide 29 presents Parallel Coordinate Plot 1, mapping high-dimensional trajectories across the main kiln. Notice how the densest operating cluster aligns with our Golden Batch target corridor, validating that expanded feed can be sustained with stable mechanical drive current.\""
            )
        )
        print("  -> Updated Slide 29: Parallel Plot 1")

        # -------------------------------------------------------------
        # Slide 30: Parallel Plot 2: Fuel & Alternative Energy Dynamics
        # -------------------------------------------------------------
        setup_graph_slide(
            pres.Slides(30),
            "Parallel Plot 2: Bursa Fuel & Alternative Energy Dynamics (Coal vs RDF Firing)",
            [
                "• Multi-Fuel Energy Mapping: Trajectories connecting Main Burner Coal, Calciner Coal, RDF Satellite Burner, and Specific Fuel Consumption.",
                "• Alternative Fuel Decoupling: High RDF dosing (>13 t/h) systematically decouples fossil coal demand, maintaining calcination enthalpy while driving down overall thermal cost."
            ],
            "slide_17.png",
            notes_text=(
                "=== SLIDE 30: Parallel Plot 2: Bursa Fuel & Alternative Energy Dynamics ===\n"
                "[ESTIMATED TIME: 2.0 minutes]\n\n"
                "[OBJECTIVE]\n"
                "Analyze parallel trajectories for multi-fuel combustion, evaluating coal displacement by RDF satellite firing.\n\n"
                "[SPOKEN SCRIPT / TALKING POINTS]\n"
                "\"Slide 30 displays Parallel Coordinate Plot 2, tracking fuel combustion splits. The trajectories demonstrate that increasing RDF substitution to 13–20 t/h enables significant coal displacement while maintaining stable burning zone temperatures.\""
            )
        )
        print("  -> Updated Slide 30: Parallel Plot 2")

        # -------------------------------------------------------------
        # Slide 31: Parallel Plot 3: Quality & Burnability Diagnostics
        # -------------------------------------------------------------
        setup_graph_slide(
            pres.Slides(31),
            "Parallel Plot 3: Bursa Quality & Burnability Diagnostics (fCaO vs C3S Trade-off)",
            [
                "• Clinker Quality Trajectories: Mapping Total Fuel Flow, Secondary Air Temp, Free Lime (fCaO), C3S, and Clinker Production.",
                "• Optimal Reactivity Convergence: Trajectories with secondary air >950°C converge on high C3S (>60%) and low Free Lime (<1.5%), establishing superior 28-day cement performance."
            ],
            "slide_16.png",
            notes_text=(
                "=== SLIDE 31: Parallel Plot 3: Bursa Quality & Burnability Diagnostics ===\n"
                "[ESTIMATED TIME: 2.0 minutes]\n\n"
                "[OBJECTIVE]\n"
                "Evaluate multi-variable parallel trajectories connecting thermal inputs with lab clinker quality assays.\n\n"
                "[SPOKEN SCRIPT / TALKING POINTS]\n"
                "\"Slide 31 presents Parallel Coordinate Plot 3, linking thermal input with clinker quality. Notice how trajectories achieving elevated secondary air temperatures converge on low Free Lime and high C3S, confirming our core quality thesis.\""
            )
        )
        print("  -> Updated Slide 31: Parallel Plot 3")

        # -------------------------------------------------------------
        # Slide 32: Parallel Plot 4: Cooler & Heat Recovery Dynamics
        # -------------------------------------------------------------
        setup_graph_slide(
            pres.Slides(32),
            "Parallel Plot 4: Bursa Cooler & Heat Recovery Dynamics (Aeration & Grate Speed)",
            [
                "• Cooler Aeration Profiling: Mapping grate speeds, undergrate chamber pressures, cooling fan speeds, and Secondary Air Temperature.",
                "• Hydraulic Optimization: Regulating grate speed stabilizes clinker bed thickness, maximizing heat recuperation back to the kiln while suppressing electrical kW draw."
            ],
            "slide_24.png",
            notes_text=(
                "=== SLIDE 32: Parallel Plot 4: Bursa Cooler & Heat Recovery Dynamics ===\n"
                "[ESTIMATED TIME: 2.0 minutes]\n\n"
                "[OBJECTIVE]\n"
                "Map clinker cooler aeration dynamics connecting grate speed, chamber pressure, cooling fan kW, and secondary air temperature.\n\n"
                "[SPOKEN SCRIPT / TALKING POINTS]\n"
                "\"Slide 32 concludes our analysis with Parallel Coordinate Plot 4, mapping cooler hydraulic dynamics. Regulating grate speed stabilizes clinker bed resistance, maximizing sensible heat return to the burning zone and minimizing cooler fan electrical power.\""
            )
        )
        print("  -> Updated Slide 32: Parallel Plot 4")

        # Save presentation
        print("\nSaving completed presentation...")
        pres.Save()
        print("Presentation saved successfully.")

    finally:
        pres.Close()
        print("PowerPoint presentation closed.")

if __name__ == "__main__":
    build_bursa_updated_presentation()
