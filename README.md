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

## Method of solution

The solver uses a classic **Euler–Bernoulli beam FEM** with **2 DOF per node**  
(vertical displacement *u<sub>y</sub>* and rotation *θ<sub>z</sub>*).

1) **Element stiffness (4×4) and global assembly**  
For each span (element length `L`, Young’s modulus `E`, inertia `I`) the local
stiffness is:
\[
k_e=\frac{EI}{L^3}
\begin{bmatrix}
12 & 6L & -12 & 6L \\
6L & 4L^2 & -6L & 2L^2 \\
-12 & -6L & 12 & -6L \\
6L & 2L^2 & -6L & 4L^2
\end{bmatrix}
\]
Code: `element__k(…)=… ; dl__k(…)+=…`  
The global stiffness matrix `dl__k` is built by assembling all `k_e`
(variables: `element__young`, `element__iz`, `element__long`).

2) **Global load vector**  
- **Point loads** at nodes go directly into `dl__f`.  
- **Uniform distributed loads** `q` on an element are converted to **equivalent
nodal forces** and added to the global RHS:
\[
f^{eq}_e = \left[-\frac{qL}{2},\;-\frac{qL^2}{12},\;-\frac{qL}{2},\;+\frac{qL^2}{12}\right]
\]
Code: `dle__feq(0..3)` → accumulated into `dl__f` and tracked per element in `element__f`.

3) **Boundary conditions (supports)**  
For every support (`noeud__appui = True`) the vertical DOF is fixed:
rows/columns of `dl__k` are zeroed, the diagonal set to 1, and the RHS set to 0.  
Code: zeroing loop on `dl__k`, then `dl__k(ii,ii)=1`, `dl__f(ii)=0`.

4) **Linear solve**  
Solve `dl__k · dl__u = dl__f` for the global displacement vector `dl__u`
using **Gaussian elimination + back substitution** (no external dependency).

5) **Element end forces & reactions**  
Recover element end forces by:
\[
f_e = k_e\,u_e + f^{eq}_e
\]
Code: loop over `element__k` and `dl__u` to update `element__f`.  
Support reactions `noeud__ry` are taken from the adjacent element end forces.

6) **Post-processing: V, M, and deflection**  
Within each element, the deflection field is reconstructed with **cubic Hermite
shape functions** from the nodal DOF (`u_y` and `θ_z`), with the distributed-load
contribution superposed. Shear `V` and moment `M` follow from the element end
forces and `q`.  
The code samples the field along each span, then extracts:
- maximum deflection per span,  
- zero-moment (inflection) points,  
- span maxima/minima of bending moment.  
Code: arrays `x__uy`, `x__mfz`, and trackers `travee__uy_max`, `travee__mfz_0`,
`travee__mfz_pos_max`, `travee__mfz_neg_min`.

**Notes**
- Mixed materials are handled natively: `E` and `I` may vary by element.  
- Units are consistent with the input sheet (SI in the template).  
- The whole pipeline is visible in the workbook: assemble → apply BCs → solve → recover → plot.


## Disclaimer

This tool is intended for educational and quick-analysis purposes.  
It is not certified for structural design or regulatory compliance.  
Always verify results against engineering standards and professional judgment.

---

## License

MIT License — free for personal and commercial use.  
Attribution appreciated.
