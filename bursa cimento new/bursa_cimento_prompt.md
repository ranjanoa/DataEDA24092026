⚠️ **AI INSTRUCTION ALERTS (READ FIRST)**:
1. **DO NOT modify, edit, or reformat any of the `[SCATTER: ...]`, `[DUALPLOT: ...]`, or `[PARALLEL: ...]` tags.** Output them exactly as written.
2. **DO NOT change the variable names** inside the tags. They deliberately use derived aliases (e.g., `KILN MAIN DRIVE (M02) CURRENT`, `SECONDARY AIR TEMP`) to avoid parsing errors. 
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
- **Kiln Main Drive Current (`KILN MAIN DRIVE (M02) CURRENT`)**: average = 179.22 A | peak = 431.73 A | Optimal Range = 130.97 - 437.42 A
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
2. **Drive Torque & Kiln Speed Coordination**: Kiln speed averages **3.11 RPM** with a drive current of **179.22 A**. Drive current spikes nearing 431.7 A indicate a heavily loaded kiln bed.
3. **Calcination & Preheater Thermal Profile**: The preheater and calciner operate with coal and RDF fuel splits. Balancing these ensures homogeneous meal decarbonation before reaching the kiln inlet.
4. **Clinker Quality & Burnability (SCAO / C3S)**: Managing clinker quality is a delicate balance. Free Lime (`KLINKER SCAO`) averages **2.03%** and C3S averages **59.18%**. High Free Lime spikes (>2.5%) indicate hard-burning meal or thermal starvation, demanding precise fuel/air adjustments to prevent under-burnt clinker without over-firing.
5. **Combustion & Alternative Fuel Utilization**: Operating with alternative fuels (RDF average **7.41 t/h**), the preheater outlet maintains high O2 levels (4.37%). Optimization of fuel splits is necessary to safely lower excess oxygen and reduce total coal consumption.
6. **Emissions Control (NOX)**: NH3 Consumption averages **145.0**, acting as the primary proxy for NOX reduction efforts, directly tied to sintering zone temperatures and combustion O2.

---

### STEP 2 — Feed Dynamics & Main Drive Stress Analysis

**Graph 1 — Total Feed Rate vs Main Drive Current**
[SCATTER: X=KILN_FEED | Y=KILN_MAIN_DRIVE_(M02)_CURRENT | COLOR=KILN_SPEED | SCALE=Jet]
*Engineering Insight (Mechanical Drive Stress)*: `KILN FEED` directly dictates the mechanical torque needed. Surges in drive current above 250 A without a corresponding increase in feed indicate unstable kiln coating, potential ring formations in the burning zone, or severe snowmen build-up forcing the mechanical drive into overload.

**Graph 2 — Main Drive Current vs Kiln Speed**
[SCATTER: X=KILN_SPEED | Y=KILN_MAIN_DRIVE_(M02)_CURRENT | COLOR=SECONDARY_AIR_TEMP | SCALE=Viridis]
*Engineering Insight (RPM Matching)*: As feed approaches 370 t/h, kiln speed must be rigidly pushed to maintain an optimal 11-13% volumetric filling degree to prevent uncalcined meal from flushing into the burning zone. This strict adherence is confirmed by the massive **+0.995 correlation between `KILN_SPEED` and `KILN_FEED`**.

---

### STEP 3 — Fuel Combustion & Specific Fuel Consumption (SFC)

**Graph 3 — Specific Fuel Consumption vs Secondary Air Temperature**
[SCATTER: X=SECONDARY_AIR_TEMP | Y=Specific_Fuel_Consumption | COLOR=Clinker_Production_tph | SCALE=Hot]
*Engineering Insight (Fuel Substitution via Recuperation)*: `SECONDARY AIR TEMP` is our most critical thermal flywheel. Operating consistently above 1000°C directly lowers the `Specific_Fuel_Consumption`, pushing thermal efficiency closer to world-class standards by minimizing the need for expensive main burner coal.

**Graph 4 — Main Burner Coal vs RDF Utilization**
[SCATTER: X=MAIN_BURNER_COAL | Y=RDF_SATELLITEBURNER | COLOR=Specific_Fuel_Consumption | SCALE=RdBu]
*Engineering Insight (Alternative Fuel Limits)*: High `RDF_SATELLITEBURNER` flow must be carefully balanced against `MAIN_BURNER_COAL`. Over-reliance on RDF without proper burner momentum risks severe CO spikes, thermal decoupling, and dragging the clinker liquid phase formation further back into the kiln.

---

### STEP 4 — Clinker Quality (Free Lime & C3S) vs Process Conditions

**Graph 5 — Free Lime (SCAO) vs Secondary Air Temperature**
[SCATTER: X=SECONDARY_AIR_TEMP | Y=KLINKER_SCAO | COLOR=KLINKER_C3S | SCALE=RdBu]
*Engineering Insight (Burnability Indicator)*: `KLINKER SCAO` (Free Lime) is the definitive indicator of burning zone health. When `SECONDARY AIR TEMP` drops and Free Lime spikes above 2.5%, the kiln is suffering from thermal starvation, leading to under-burned, dusty clinker that will degrade final cement strength.

**Graph 6 — C3S vs Free Lime (SCAO)**
[SCATTER: X=KLINKER_SCAO | Y=KLINKER_C3S | COLOR=Specific_Fuel_Consumption | SCALE=Viridis]
*Engineering Insight (Quality vs Efficiency Trade-off)*: High C3S requires intense, localized heat. The operational challenge is to hit the 'Golden Zone' (SCAO ~1.0-1.5%, C3S > 60%) using maximized secondary air heat recuperation rather than just dumping more `MAIN_BURNER_COAL`, which unnecessarily penalizes the SFC and risks burning out the refractory brick lining.

---

### STEP 5 — Emissions Stability & NH3 Control

**Graph 7 — NH3 Consumption vs Secondary Air Temperature**
[SCATTER: X=SECONDARY_AIR_TEMP | Y=NH3_CONSUMPTION | COLOR=PRE_HEATER_OUTLET_O2 | SCALE=Viridis]
*Engineering Insight (NOX Control via NH3)*: Thermal NOX generation is highly sensitive to peak flame temperature and excess oxygen. High `SECONDARY AIR TEMP` combined with excess `PRE HEATER OUTLET O2` triggers rampant NOX formation, necessitating heavy NH3 dosing. The operator must choke back excess air to minimize NOX at the source, preventing excessive reagent (`NH3_CONSUMPTION`) costs.

---

## PART 2 — EXHAUST, PREHEATER & COOLER DYNAMICS

### STEP 6 — Draft Symmetry & ID Fan Balance

**Graph 8 — Preheater Outlet O2 vs Preheater Fan Outlet**
[SCATTER: X=PRE_HEATER_OUTLET_O2 | Y=PRE_HEATER_FAN_OUTLET | COLOR=KILN_FEED | SCALE=Jet]
*Engineering Insight (Aerodynamics & Draft Symmetry)*: Proper draft is non-negotiable for a 340+ t/h line. The system shows near-perfect symmetry with a **+0.998 correlation between `PRESSURE_AFTER_PRE-HEATER_1` and `2`**. Any deviation from this symmetry immediately flags cyclone blockages or massive false air ingress. We must ruthlessly minimize system pressure drop to maintain production without red-lining the ID fan capacity.

---

### STEP 7 — Clinker Cooler Bed Resistance & Secondary Air

**Graph 9 — Cooler Exhaust Temp vs Secondary Air Temp**
[SCATTER: X=COOLER_EXHAUST_AIR_TEMP | Y=SECONDARY_AIR_TEMP | COLOR=COOLER_FAN_KW | SCALE=Hot]
*Engineering Insight (Fan Power Optimization)*: The clinker bed depth dictates the entire cooler thermal efficiency. Deep beds force cooling fans into severe power draw (proven by the **+0.991 correlation between `COOLER_FAN_SPEED` and `COOLER_FAN_KW`**). Speeding up the grate reduces pressure and sheds massive electrical kW load, but the operator must execute this without blowing cold air through the bed and destroying our `SECONDARY_AIR_TEMP` thermal flywheel.

---

### STEP 8 — 3D System Operating Envelopes

**3D Quality & Thermal Space:**
[SCATTER3D: X=SECONDARY_AIR_TEMP | Y=KLINKER_SCAO | Z=KLINKER_C3S | COLOR=Specific_Fuel_Consumption | SCALE=Jet]

**3D Cooler & Fuel Efficiency Space:**
[SCATTER3D: X=RDF_SATELLITEBURNER | Y=Total_Coal_Flow | Z=SECONDARY_AIR_TEMP | COLOR=Clinker_Production_tph | SCALE=Viridis]

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
[PARALLEL: Total_Fuel_Flow, SECONDARY_AIR_TEMP, KLINKER_SCAO, KLINKER_C3S, Clinker_Production_tph | COLOR: KLINKER_SCAO]

2. **Bursa Main Kiln Process Dynamics**:
[PARALLEL: KILN_FEED, Clinker_Production_tph, KILN_SPEED, KILN_MAIN_DRIVE_(M02)_CURRENT, Specific_Fuel_Consumption, PRE_HEATER_OUTLET_O2 | COLOR: SECONDARY_AIR_TEMP]

3. **Bursa Fuel & Alternative Energy Dynamics**:
[PARALLEL: KILN_FEED, MAIN_BURNER_COAL, CALCINER_COAL, RDF_SATELLITEBURNER, Specific_Fuel_Consumption | COLOR: Clinker_Production_tph]

4. **Bursa Cooler & Heat Recovery Dynamics**:
[PARALLEL: GRATE_SPEED, COOLER_FAN_SPEED, COOLER_FAN_KW, COOLER_EXHAUST_AIR_TEMP, SECONDARY_AIR_TEMP | COLOR: SECONDARY_AIR_TEMP]
