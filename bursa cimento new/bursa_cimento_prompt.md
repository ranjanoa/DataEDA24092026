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

### Dataset Overview:
- **Resolution**: 1-minute interval data
- **Size**: 240,414 historical process records (approx. 167 days of continuous operation from February 11, 2026 to August 6, 2026)
- **Variables**: 32 synchronized pyro-process & laboratory quality variables

### Process Statistics:
- **Total Kiln Feed (`KILN FEED`)**: average = 340.9 t/h | peak = 406.3 t/h | Stable Range = 0 - 350.1 t/h
- **Clinker Production (`Clinker_Production_tph`)**: average = 221.6 t/h (Derived at 0.65 ratio)
- **Kiln Speed (`KILN SPEED`)**: average = 3.11 rpm | High-Feed Stable Range (at >340 t/h) = 3.22 - 3.50 rpm
- **Kiln Main Drive Current (`KILN MAIN DRIVE (M02) CURRENT`)**: average = 179.22 A | High-Feed Stable Range = 295.47 - 318.40 A
- **Main Burner Coal (`MAIN BURNER COAL`)**: average = 8.90 t/h | High-Feed Stable Range = 8.81 - 9.71 t/h
- **Calciner Coal (`CALCINER COAL`)**: average = 11.49 t/h | High-Feed Stable Range = 8.88 - 16.04 t/h
- **Alternative Fuel (`RDF SATELLITEBURNER`)**: average = 7.41 t/h | High-Feed Stable Range = 0.00 - 13.93 t/h
- **Secondary Air Temp (`SECONDARY AIR TEMP`)**: average = 811.5 °C | High-Feed Stable Range = 698.95 - 921.88 °C
- **Tertiary Air Temp (`TERTIARY AIR TEMP`)**: average = 606.79 °C | High-Feed Stable Range = 939.27 - 993.36 °C
- **Sintering Zone Temp (`SINTERING ZONE TEMP`)**: average = 616.74 °C | High-Feed Stable Range = 1017.24 - 1079.05 °C
- **Calciner Outlet Temp (`CALCINER OUTLET A TEMP`)**: average = 551.08 °C | High-Feed Stable Range = 882.01 - 890.09 °C
- **Kiln Inlet Temp (`KILN INLET TEMP`)**: average = 1281.31 °C | High-Feed Stable Range = 803.82 - 1091.99 °C
- **Cooler Exhaust Temp (`COOLER EXHAUST AIR TEMP`)**: average = 85.57 °C | High-Feed Stable Range = 131.41 - 142.84 °C
- **Pre Heater Outlet O2 (`PRE HEATER OUTLET O2`)**: average = 4.37% | High-Feed Stable Range = 3.10 - 4.46%
- **Pre Heater Outlet CO (`PRE HEATER OUTLET CO`)**: average = 0.043% | High-Feed Stable Range = 0.02 - 0.05%
- **NOX Proxy (`NH3 CONSUMPTION`)**: average = 145.00 | High-Feed Stable Range = 1.0 - 168.9
- **Clinker C3S (`KLINKER C3S`)**: average = 59.18% | High-Feed Stable Range = 59.60 - 63.90%
- **Free Lime / SCAO (`KLINKER SCAO`)**: average = 2.03% | High-Feed Stable Range = 0.98 - 1.94%
- **Pre Heater Fan Outlet Draft (`PRE HEATER FAN OUTLET`)**: High-Feed Stable Range = -4.95 to -3.32
- **Calciner O2 A (`CALCINER O2 A`)**: High-Feed Stable Range = 2.66 - 4.50%
- **Calciner Outlet A Pressure (`CALCINER OUTLET A PRESSURE`)**: High-Feed Stable Range = -15.19 to -12.69 mbar
- **Kiln Inlet Pressure (`KILN INLET PRESSURE`)**: High-Feed Stable Range = -1.65 to -1.03 mbar
- **Kiln Inlet O2 (`KILN INLET O2`)**: High-Feed Stable Range = 2.65 - 3.86%
- **Pressure After Pre-Heater 1 (`PRESSURE AFTER PRE-HEATER 1`)**: High-Feed Stable Range = -37.83 to -33.89 mbar
- **Pressure After Pre-Heater 2 (`PRESSURE AFTER PRE-HEATER 2`)**: High-Feed Stable Range = -37.88 to -33.80 mbar
- **Cooler Fan Power (`COOLER FAN KW`)**: High-Feed Stable Range = 340.76 - 402.82 kW
- **Cooler Fan Speed (`COOLER FAN SPEED`)**: High-Feed Stable Range = 67.11 - 75.26 rpm
- **Cooler Chamber Pressure (`CHAMBER PRESSURE_COOLER FAN DRAFT`)**: High-Feed Stable Range = -0.74 to -0.60 mbar
- **Grate Speed (`GRATE SPEED`)**: High-Feed Stable Range = 4.10 - 4.99 rpm

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
State the complete operating window as a filled table grounded in proven high-capacity operation (>340 t/h feed):
| Variable | Golden Batch Lower | Golden Batch Upper | Current Mean | Gap |
|---|---|---|---|---|
| Clinker_Production_tph | 240.00 | 260.00 | 221.58 | +18.42 to +38.42 t/h |
| KILN_FEED | 369.00 | 400.00 | 340.90 | +28.10 to +59.10 t/h |
| KILN_SPEED | 3.22 | 3.50 | 3.11 | +0.11 to +0.39 rpm |
| KILN_MAIN_DRIVE_(M02)_CURRENT | 295.47 | 318.40 | 179.22 | +116.25 to +139.18 A |
| MAIN_BURNER_COAL | 8.81 | 9.71 | 8.90 | near optimal |
| CALCINER_COAL | 8.88 | 16.04 | 11.49 | near optimal |
| RDF_SATELLITEBURNER| 10.00 | 13.93 | 7.41 | +2.59 to +6.52 t/h |
| PRE_HEATER_OUTLET_O2 | 3.10 | 4.46 | 4.37 | near optimal |
| PRE_HEATER_OUTLET_CO | 0.02 | 0.05 | 0.04 | near optimal |
| SECONDARY_AIR_TEMP | 850.00 | 921.88 | 811.54 | +38.46 to +110.34°C |
| TERTIARY_AIR_TEMP | 939.27 | 993.36 | 606.79 | +332.48 to +386.57°C |
| SINTERING_ZONE_TEMP | 1017.24 | 1079.05 | 616.74 | +400.50 to +462.31°C |
| CALCINER_OUTLET_A_TEMP | 882.01 | 890.09 | 551.08 | +330.93 to +339.01°C |
| KILN_INLET_TEMP | 803.82 | 1091.99 | 1281.31 | -477.49 to -189.32°C |
| COOLER_EXHAUST_AIR_TEMP | 131.41 | 142.84 | 85.57 | +45.84 to +57.27°C |
| NH3_CONSUMPTION | 1.00 | 168.90 | 145.00 | near optimal |
| KLINKER_SCAO | 0.98 | 1.50 | 2.03 | -0.53 to -1.05% |
| KLINKER_C3S | 59.60 | 63.90 | 59.18 | +0.42 to +4.72% |

### STEP 10 — VALUE LOST QUANTIFICATION
Recalculate all numbers dynamically using the statistics above to highlight the percentage efficiency gains:
- **Production Opportunity**: Gap = (240.00 − 221.58) = 18.42 t/h increase. This represents an **~8.31% increase** in total clinker production throughput.
- **Alternative Fuel Savings**: Current average total coal is 20.39 t/h. The Golden Batch targets 17.69 t/h (8.81 + 8.88) by substituting with RDF. Gap = (20.39 - 17.69) = 2.70 t/h of coal saved. This represents an **~13.24% reduction** in total coal consumption.
- **Quality Cost Reduction**: Value of stabilizing `KLINKER_SCAO` below 1.5%, avoiding corrective grinding.

Present this completed summary table:
| Opportunity | Physical Quantity | Percentage Gain |
|---|---|---|
| Production Increase | 18.42 t/h extra Clinker | **+8.31%** Throughput |
| Coal Reduction (via RDF) | 2.70 t/h of Coal Saved | **-13.24%** Coal Usage |
| Quality & C3S improvement | Stable Free Lime < 1.5% | Avoided quality adjustments |
| **App Realization Potential (20% capture)** | **3.68 t/h extra Clinker, 0.54 t/h Coal Saved** | **+1.66% Throughput, -2.65% Coal** |

---

## PART 3 — TIME-SERIES TRANSIENT STABILITY

### STEP 11 — Production vs Fuel Efficiency Time Trend
**Graph 10 — Kiln Feed & Clinker vs SFC**
[DUALPLOT: KILN_FEED, Clinker_Production_tph | Specific_Fuel_Consumption]
*Engineering Insight (Transient Instability)*: Analyze the macro time-series view. Look for periods where high Kiln Feed correctly drops the SFC versus periods where SFC spikes erratically despite steady feed, indicating severe burner/cooler instability or raw mix burnability issues.

### STEP 12 — Thermal Recuperation vs Clinker Quality
**Graph 11 — Secondary Air Temp vs Free Lime (SCAO)**
[DUALPLOT: SECONDARY_AIR_TEMP | KLINKER_SCAO]
*Engineering Insight (Thermal Decoupling)*: This timeline reveals thermal decoupling. When Secondary Air Temp plunges, does the Free Lime spike immediately, or is there a delay? Prolonged periods of low temp and high SCAO confirm thermal starvation in the burning zone.

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
