⚠️ **AI INSTRUCTION ALERTS (READ FIRST)**:
1. **DO NOT modify, edit, or reformat any of the `[SCATTER: ...]`, `[DUALPLOT: ...]`, or `[PARALLEL: ...]` tags.** Output them exactly as written.
2. **DO NOT change the variable names** inside the tags. They deliberately use derived aliases (e.g., `KILN MAIN DRIVE (M01) CURRENT`, `SECONDARY AIR TEMP`) to avoid parsing errors. 
3. Output the `[DERIVED: ...]` lines exactly as they are at the top.

# BURSA CIMENTO — PYRO PROCESS OPTIMIZATION & ENGINEERING DIAGNOSTICS REPORT
*Prepared by: Process Engineering Dept, Bursa Cimento*

[DERIVED: Clinker_Production_tph = KILN_FEED * 0.65]
[DERIVED: Total_Coal_Flow = MAIN_BURNER_COAL + CALCINER_COAL]
[DERIVED: Total_Fuel_Flow = Total_Coal_Flow + RDF_SATELLITEBURNER]
[DERIVED: Specific_Fuel_Consumption = Total_Fuel_Flow / Clinker_Production_tph]

## 📊 DATASET REFERENCE STATISTICS (BURSA CIMENTO OPERATIONAL DATA)

- **Total Kiln Feed (`KILN FEED`)**: average = 340.9 t/h | peak = 406.3 t/h | Stable Range = 0 - 350.1 t/h
- **Clinker Production (`Clinker_Production_tph`)**: average = 221.6 t/h (Derived at 0.65 ratio)
- **Kiln Speed (`KILN SPEED`)**: average = 3.11 rpm | peak = 3.95 rpm | Optimal Range = 1.10 - 4.40 rpm
- **Kiln Main Drive Current (`KILN MAIN DRIVE (M01) CURRENT`)**: average = 54.38 A | peak = 71.51 A
- **Main Burner Coal (`MAIN BURNER COAL`)**: average = 8.90 t/h | Optimal Range = 3.70 - 12.71 t/h
- **Calciner Coal (`CALCINER COAL`)**: average = 11.49 t/h | Optimal Range = 0.64 - 13.64 t/h
- **Alternative Fuel (`RDF SATELLITEBURNER`)**: average = 7.41 t/h | peak = 47.63 t/h
- **Secondary Air Temp (`SECONDARY AIR TEMP`)**: average = 811.5 °C | Stable Range = 785 - 1800 °C
- **Pre Heater Outlet O2 (`PRE HEATER OUTLET O2`)**: average = 4.37% | Stable Range = 3.47 - 20.83%
- **Pre Heater Outlet CO (`PRE HEATER OUTLET CO`)**: average = 0.043% | Optimal Range = 0.02 - 0.05%
- **NOX Proxy (`NH3 CONSUMPTION`)**: average = 145.00 | Stable Range = 1.0 - 100.2
- **Clinker C3S (`KLINKER C3S`)**: average = 59.18% | Optimal Range = 57.05 - 62.95%
- **Free Lime / SCAO (`KLINKER SCAO`)**: average = 2.03% | Optimal Range = 1.17 - 2.45%

### Key System Correlation Coefficients:
- `PRESSURE_AFTER_PRE-HEATER_1` vs `PRESSURE_AFTER_PRE-HEATER_2`: **+0.998**
- `KILN_SPEED` vs `KILN_FEED`: **+0.995**
- `COOLER_FAN_SPEED` vs `COOLER_FAN_KW`: **+0.991**
- `MAIN_BURNER_COAL` vs `KILN_FEED`: **+0.983**
- `MAIN_BURNER_COAL` vs `KILN_SPEED`: **+0.981**
- `CALCINER_OUTLET_A_TEMP` vs `SINTERING_ZONE_TEMP`: **+0.982**
- `COOLING_OUTLET_TEMP` vs `COOLER_EXHAUST_AIR_TEMP`: **+0.981**

---

# PART 1 — KILN FEED DYNAMICS, THERMAL CONTROL & BURNING ZONE

### STEP 1 — Executive Summary & Operational Overview
1. **High-Capacity Volumetric Loading**: The Bursa kiln operates at an average of **340.9 t/h** total feed, sustaining peaks up to 406.3 t/h. This high throughput requires stringent thermal management.
2. **Drive Torque & Kiln Speed Coordination**: Kiln speed averages **3.11 RPM** with a drive current of **54.38 A**. Drive current spikes nearing 71.5 A indicate a heavily loaded kiln bed.
3. **Calcination & Preheater Thermal Profile**: The preheater and calciner operate with coal and RDF fuel splits. Balancing these ensures homogeneous meal decarbonation before reaching the kiln inlet.
4. **Clinker Quality & Burnability (SCAO / C3S)**: Managing clinker quality is a delicate balance. Free Lime (`KLINKER SCAO`) averages **2.03%** and C3S averages **59.18%**. High Free Lime spikes (>2.5%) indicate hard-burning meal or thermal starvation, demanding precise fuel/air adjustments to prevent under-burnt clinker without over-firing.
5. **Combustion & Alternative Fuel Utilization**: Operating with alternative fuels (RDF average **7.41 t/h**), the preheater outlet maintains high O2 levels (4.37%). Optimization of fuel splits is necessary to safely lower excess oxygen and reduce total coal consumption.
6. **Emissions Control (NOX)**: NH3 Consumption averages **145.0**, acting as the primary proxy for NOX reduction efforts, directly tied to sintering zone temperatures and combustion O2.

---

### STEP 2 — Feed Dynamics & Main Drive Stress Analysis

**Graph 1 — Total Feed Rate vs Main Drive Current**
[SCATTER: X=KILN FEED | Y=KILN MAIN DRIVE (M01) CURRENT | COLOR=KILN SPEED | SCALE=Jet]

**Graph 2 — Main Drive Current vs Kiln Speed**
[SCATTER: X=KILN SPEED | Y=KILN MAIN DRIVE (M01) CURRENT | COLOR=SECONDARY AIR TEMP | SCALE=Viridis]

**Engineering Diagnostics & Insights**:
- **Mechanical Drive Stress**: `KILN FEED` directly dictates the mechanical torque needed. Surges in drive current above 60 A without a corresponding increase in feed indicate unstable kiln coating.
- **RPM Matching**: As feed approaches 370 t/h, kiln speed must be coordinated to maintain an optimal volumetric filling degree and prevent material flushing. This strict adherence is confirmed by the massive **+0.995 correlation between `KILN_SPEED` and `KILN_FEED`**.

---

### STEP 3 — Fuel Combustion & Specific Fuel Consumption (SFC)

**Graph 3 — Specific Fuel Consumption vs Secondary Air Temperature**
[SCATTER: X=SECONDARY AIR TEMP | Y=Specific_Fuel_Consumption | COLOR=Clinker_Production_tph | SCALE=Hot]

**Graph 4 — Main Burner Coal vs RDF Utilization**
[SCATTER: X=MAIN BURNER COAL | Y=RDF SATELLITEBURNER | COLOR=Specific_Fuel_Consumption | SCALE=RdBu]

**Engineering Diagnostics & Insights**:
- **Fuel Substitution via Recuperation**: `SECONDARY AIR TEMP` is a critical efficiency metric. Operating consistently at higher temperatures lowers the `Specific_Fuel_Consumption`.
- **Alternative Fuel Limits**: High `RDF SATELLITEBURNER` flow must be balanced against `MAIN BURNER COAL` to prevent unstable burning zones.

---

### STEP 4 — Clinker Quality (Free Lime & C3S) vs Process Conditions

**Graph 5 — Free Lime (SCAO) vs Secondary Air Temperature**
[SCATTER: X=SECONDARY AIR TEMP | Y=KLINKER SCAO | COLOR=KLINKER C3S | SCALE=RdBu]

**Graph 6 — C3S vs Free Lime (SCAO)**
[SCATTER: X=KLINKER SCAO | Y=KLINKER C3S | COLOR=Specific_Fuel_Consumption | SCALE=Viridis]

**Engineering Diagnostics & Insights**:
- **Burnability Indicator**: `KLINKER SCAO` (Free Lime) is the definitive indicator of burning zone health. When `SECONDARY AIR TEMP` drops, Free Lime spikes above 2.5%, indicating under-burning and poor nodulization.
- **Quality vs Efficiency Trade-off**: High C3S requires intense heat, which naturally lowers Free Lime but can increase `Specific_Fuel_Consumption`. The goal is to hit the sweet spot (SCAO ~1.0-1.5%, C3S > 60%) using recuperated heat rather than just pushing more `MAIN BURNER COAL`.

---

### STEP 5 — Emissions Stability & NH3 Control

**Graph 7 — NH3 Consumption vs Secondary Air Temperature**
[SCATTER: X=SECONDARY AIR TEMP | Y=NH3 CONSUMPTION | COLOR=PRE HEATER OUTLET O2 | SCALE=Viridis]

**Engineering Diagnostics & Insights**:
- **NOX Control via NH3**: High `SECONDARY AIR TEMP` and high `PRE HEATER OUTLET O2` often trigger higher NOX formation, necessitating more NH3. Balancing excess air is key to minimizing both emissions and reagent costs.

---

## PART 2 — EXHAUST, PREHEATER & COOLER DYNAMICS

### STEP 6 — Draft Symmetry & ID Fan Balance

**Graph 8 — Preheater Outlet O2 vs Preheater Fan Outlet**
[SCATTER: X=PRE HEATER OUTLET O2 | Y=PRE HEATER FAN OUTLET | COLOR=KILN FEED | SCALE=Jet]

**Engineering Diagnostics & Insights**:
- **Aerodynamics & Draft Symmetry**: Proper draft is non-negotiable. The system shows near-perfect symmetry with a **+0.998 correlation between `PRESSURE_AFTER_PRE-HEATER_1` and `2`**. We must minimize system pressure drop and false air to maintain production without hitting the fan ceiling.

---

### STEP 7 — Clinker Cooler Bed Resistance & Secondary Air

**Graph 9 — Cooler Exhaust Temp vs Secondary Air Temp**
[SCATTER: X=COOLER EXHAUST AIR TEMP | Y=SECONDARY AIR TEMP | COLOR=COOLER FAN KW | SCALE=Hot]

**Engineering Diagnostics & Insights**:
- **Fan Power Optimization**: Deep clinker beds force cooling fans into high power draw (proven by the **+0.991 correlation between `COOLER_FAN_SPEED` and `COOLER_FAN_KW`**). Optimizing the grate reduces pressure and saves kW, but must be carefully balanced against preserving high `SECONDARY AIR TEMP` for the burning zone.

---

### STEP 8 — 3D System Operating Envelopes

**3D Quality & Thermal Space:**
[SCATTER3D: X=SECONDARY AIR TEMP | Y=KLINKER SCAO | Z=KLINKER C3S | COLOR=Specific_Fuel_Consumption | SCALE=Jet]

**3D Cooler & Fuel Efficiency Space:**
[SCATTER3D: X=RDF SATELLITEBURNER | Y=Total_Coal_Flow | Z=SECONDARY AIR TEMP | COLOR=Clinker_Production_tph | SCALE=Viridis]

**3D Operating Cluster Interpretation**:
The 3D envelopes map the "Bursa Golden Zone" — the exact combination of Fuel Split (RDF vs Coal) and Recuperation (Secondary Air Temp) that yields maximum clinker quality (C3S > 60%, SCAO < 1.5%) at minimum thermal load.

---

### STEP 9 — Golden Batch Complete Definition
State the complete operating window as a filled table:
| Variable | Golden Batch Lower | Golden Batch Upper | Current Mean | Gap |
|---|---|---|---|---|
| Clinker_Production_tph | 240.00 | 260.00 | 221.58 | +18.42 to +38.42 t/h |
| KILN_FEED | 369.00 | 400.00 | 340.90 | +28.10 to +59.10 t/h |
| MAIN_BURNER_COAL | 8.00 | 9.00 | 8.90 | near optimal |
| CALCINER_COAL | 10.00 | 13.00 | 11.49 | near optimal |
| RDF_SATELLITEBURNER| 13.00 | 20.00 | 7.41 | +5.59 to +12.59 t/h |
| PRE_HEATER_OUTLET_O2 | 3.50 | 4.50 | 4.37 | near optimal |
| SECONDARY_AIR_TEMP | 950.00 | 1100.00 | 811.54 | +138.46 to +288.46°C |
| KLINKER_SCAO | 1.17 | 1.50 | 2.03 | -0.53 to -0.86% |
| KLINKER_C3S | 60.00 | 63.00 | 59.18 | +0.82 to +3.82% |

### STEP 10 — VALUE LOST QUANTIFICATION
Recalculate all numbers dynamically using the statistics above:
- **Production Opportunity**: Gap = (240.00 − 221.58) × 8,120 hours/year = 149,606 t/yr. Calculate financial value at **$50/ton** of clinker ($7,480,300/yr).
- **Alternative Fuel Savings**: Calculate savings of substituting coal with higher RDF usage.
- **Quality Cost Reduction**: Value of stabilizing `KLINKER_SCAO` below 1.5%, avoiding corrective grinding.

Present this completed summary table:
| Opportunity | Annual Quantity | Financial Value |
|---|---|---|
| Production (conservative) | 149,606 t/yr | ~$7.48M at $50/t |
| Fuel & RDF saving | [Calculated t/yr] | [Calculated $ value] |
| Quality & C3S improvement | Better mineralogy | Avoided quality adjustments |
| **TOTAL (conservative)** | | **[Sum of Financial Value]** |

---

## ⚡ BONUS — PARALLEL PLOTS FOR EXPORT

1. **Bursa Quality & Burnability Diagnostics**:
[PARALLEL: Total_Fuel_Flow, SECONDARY AIR TEMP, KLINKER SCAO, KLINKER C3S, Clinker_Production_tph | COLOR: KLINKER SCAO]

2. **Bursa Main Kiln Process Dynamics**:
[PARALLEL: KILN FEED, Clinker_Production_tph, KILN SPEED, KILN MAIN DRIVE (M01) CURRENT, Specific_Fuel_Consumption, PRE HEATER OUTLET O2 | COLOR: SECONDARY AIR TEMP]

3. **Bursa Fuel & Alternative Energy Dynamics**:
[PARALLEL: KILN FEED, MAIN BURNER COAL, CALCINER COAL, RDF SATELLITEBURNER, Specific_Fuel_Consumption | COLOR: Clinker_Production_tph]

4. **Bursa Cooler & Heat Recovery Dynamics**:
[PARALLEL: GRATE SPEED, COOLER FAN SPEED, COOLER FAN KW, COOLER EXHAUST AIR TEMP, SECONDARY AIR TEMP | COLOR: SECONDARY AIR TEMP]
