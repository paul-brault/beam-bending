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
