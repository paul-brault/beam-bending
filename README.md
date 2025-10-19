# Beam-Bending
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
![Excel](https://img.shields.io/badge/Excel-365%2B-green)
![VBA](https://img.shields.io/badge/VBA-compatible-blue)

Open Excel-based beam calculator — unlimited supports, multi-material spans, instant shear/moment/deflection charts.  
Contact: beam.bending@gmail.com — Paul Brault

---

## Overview

Beam-Bending is a lightweight Excel/VBA tool for quick beam analysis.  
It runs entirely inside Excel — no external dependencies, no UI layers, and no double data entry.  
Supports unlimited spans and materials, with live charts for shear (V), moment (M), and deflection.

---

## Screenshots

### Input sheet
Beam geometry, materials, and loads are entered directly in Excel cells.

![Input sheet](docs/input_UI.png)

### Output sheet
Instant shear, bending moment, and deflection diagrams are generated automatically.

![Output sheet](docs/output_UI.png)

---

## Features

- Unlimited supports and spans  
- Mixed materials per span (variable E and I)  
- Point and distributed loads  
- Instant V/M/deflection charts  
- 100% visible Excel logic (no hidden macros or hidden sheets)  
- Compatible with Excel 365 (or older versions with minor adjustments on formulas, VBA is already compatible)

---

## Installation & Usage

1. Download `Beam-Bending.xlsm`  
2. Enable macros when prompted  
3. Enter supports, spans, material properties (E, I), and loads  
4. View automatic chart updates for shear, moment, and deflection  

---

## Input Data Format

The solver expects one line of input containing **9 fields**, separated by `;`.  
Inside each field, values are separated by `:`.  
Units must be consistent (SI).

### Field Definitions

1. **Support positions (x) [m]**  
   Beam span supports.  
   - If no overhangs (no cantilevers), include `0` as the first value.  
   - Example: `0:3:7`

2. **Span end positions (x) [m]**  
   End of each continuous beam section.  
   - If only one material/geometry, provide just the total length.  
   - Example: `10`

3. **Young’s modulus per span [N/m²]**  
   One value per span section (same count as span ends).  
   - Example: `2.1E11`

4. **Moment of inertia per span [m⁴]**  
   One value per span section (same count as span ends).  
   - Example: `8.5E-6`

5. **Point load positions (x) [m]**  
   Positions of nodal loads.  
   - Example: `4:8`

6. **Point load magnitudes [N]**  
   Values corresponding to the positions above (same count).  
   - Example: `-5000:-3000`

7. **Distributed load start positions (x) [m]**  
   One value for each distributed load.  
   - Example: `2`

8. **Distributed load end positions (x) [m]**  
   Same count as distributed load starts.  
   - Example: `5`

9. **Distributed load intensities [N/m]**  
   Constant intensity, same count as distributed load starts.  
   - Example: `-2000`

### Rules

- Use `:` inside fields, `;` between fields.  
- Beam end positions must be strictly increasing.  
- Number of values must match across related fields (e.g., point load positions ↔ point load magnitudes).  
- At least **2 supports** are required.  
- At least **one nonzero load** is required.  

---

## Example with RDM6 / RDM7

The main workbook (`beam-bending.xlsm`) already includes a preloaded example reproducing the same case analyzed in **RDM6/RDM7**.

This benchmark illustrates:
- A continuous beam with **6 supports** and **7 spans**  
- **Three different materials** (varying E and I across spans)  
- A combination of **point** and **uniformly distributed loads**  
- Instant computation of reactions, bending moments, and deflections  
- Perfect consistency with the results from RDM6/RDM7  

This comparison validates the accuracy and speed of the Excel/VBA solver.

### Visualization

![RDM6/RDM7 comparison](docs/example_rdm6-rdm7.png)

Reference model: `example_rdm6-rdm7.fle`

---

## Market Applications

This Excel-based structural beam calculator is designed for projects where **speed, volume, and efficiency** are critical.  
Unlike heavier software or Python scripts, the **VBA backend is optimized for Windows**, allowing hundreds of beam simulations in just fractions of a second — with **no duplicate data entry** since everything is contained directly in the spreadsheet environment.  
It supports **flexion (bending) cases**, **multiple beams**, and **multiple load conditions** in a single run, bringing advanced structural analysis into a lightweight tool engineers already use daily.

### Warehouse Racking and Storage Systems
High-volume design of **pallet racks, warehouse shelving, flow racks, AS/RS systems, and mezzanines**.  
Used for **batch calculations of rack beams**, checking **deflection, bending stress, and load capacity** according to **EN 15512, RMI, and AISC standards**.

### Solar Panel Mounting Structures
Rapid simulation of **solar racking systems, rooftop PV mounts, and ground-mount frames**.  
Handles **wind load, snow load, and aluminum profile bending checks** under **Eurocode 1, Eurocode 9, and ASCE 7**.  
Ideal for **iterating many design variants quickly** during solar structural engineering projects.

### Mechanical Supports and Strut Channel Systems
Fast verification of **HVAC supports, pipe racks, strut channel frames, and MEP modular structures**.  
Automates **span checks, bending resistance, and deflection limits** for large sets of support elements using **Eurocode 3, AISI, and ASTM**.  
Eliminates manual re-entry by keeping data fully inside Excel.

### Scaffolding and Formwork
Optimized for **batch checking of scaffold ledgers, braces, temporary beams, and formwork systems**.  
Ensures **structural safety under multiple load cases** with **EN 12811, OSHA, and BS 5975 compliance**.  
Well-suited for **scaffolding contractors running repetitive beam verifications**.

### Mezzanines and Industrial Platforms
Efficient sizing of **free-standing mezzanines, industrial work platforms, and warehouse decks**.  
Performs **beam and column bending analysis, deflection checks, and load capacity verifications** according to **AISC 360 and IBC**.  
Accelerates **large-batch design studies** for platform layouts.

### Modular Aluminum Profiles and Machine Frames
Structural checks for **aluminum extrusion frames, conveyors, machine bases, and automation structures**.  
Excel-driven **deflection and bending analysis** using **Eurocode 9 and DIN 4113**.  
Useful for **running many profile span variations in seconds** without needing external CAD/FEA.

### Guardrails and Road Safety Barriers
High-speed analysis of **roadside guardrails, crash barriers, and bridge parapets**.  
Verifies **beam deflection and impact resistance** according to **EN 1317 and AASHTO LRFD**.  
Enables **rapid testing of multiple barrier configurations** for safety compliance.

### Key Benefits
- **Batch simulations:** run dozens or hundreds of cases in a fraction of a second.  
- **Optimized VBA engine:** faster than Python for spreadsheet-native workflows, with direct Windows integration.  
- **No double entry:** all inputs, outputs, and results live in Excel.  
- **Multiple beams and load cases:** flexible for real-world engineering projects.  
- **Advanced bending analysis in a lightweight tool:** bridges the gap between manual spreadsheets and heavy FEA software.

---

## Method of Resolution

The beam solver is based on the **Euler–Bernoulli finite element method (FEM)** for linear static analysis.

1. **Discretization**  
   The continuous beam is divided into finite elements, each connecting two nodes.  
   Each node carries two degrees of freedom: vertical displacement and rotation.  
   This allows accurate modeling of deflection and bending behavior, even across spans with different materials.

2. **Element stiffness formulation**  
   For each element, a local 4×4 stiffness matrix is derived from the Euler–Bernoulli beam equations,  
   relating nodal forces and displacements through the material properties (E, I) and the element length (L).  
   The local stiffness matrices are then assembled into a global stiffness matrix that represents the entire structure.

3. **Loading**  
   Point loads are applied directly at the nodes, while distributed loads are converted into equivalent nodal forces.  
   These contributions are combined into a single global load vector.

4. **Boundary conditions**  
   Supports (simple or fixed) are introduced by constraining the relevant degrees of freedom.  
   This ensures that the structure reacts consistently with the physical boundary conditions of the beam.

5. **System resolution**  
   The global linear system (stiffness matrix × displacements = loads) is solved using Gaussian elimination.  
   The resulting displacements and rotations at each node define the deformation of the entire beam.

6. **Post-processing**  
   Once the nodal displacements are known, internal forces and moments are computed within each element.  
   Shear force, bending moment, and deflection diagrams are then reconstructed along the beam length  
   using cubic Hermite shape functions.  
   The solver automatically extracts the maximum deflections, support reactions, and bending moment extrema for each span.

This classical finite element formulation provides both accuracy and computational efficiency,  
while remaining transparent and fully accessible inside Excel.


## Disclaimer

This tool is intended for educational and quick-analysis purposes.  
It is not certified for structural design or regulatory compliance.  
Always verify results against engineering standards and professional judgment.

---

## License

MIT License — free for personal and commercial use.  
Attribution appreciated.
