# GILL: General Insurance Lambda Library

[![license](https://img.shields.io/badge/license-GPL%20v3-blue.svg)](https://github.com/BayesOnATrueStory/GILL/blob/main/LICENSE)

GILL (pronounced: 'gill') is a library of 14 general-purpose general insurance (or P&C) functions meant for actuaries using Excel 2024 or beyond. These functions are wrapped into a .bas package that Excel users can import and use in their workflows.

The intent is to cover common or useful functions within P&C pricing, reserving and reinsurance work. The intent is to have this lightweight batch traveling with the workbook, in that no conversion to .xlsm is needed, nothing additional needs to be installed next, and no add-in needs to be kept loaded, as long as the file's saved once the functions are first installed.

------------------------------------------------------------------------

### Why These 14 Functions?

I picked these functions with 2 criteria in mind:

-   **Is it a VBA replacement?** The advent of `LAMBDA` in Excel has allowed for multi-step calculations (requiring memory, like running totals) that can now be carried out in single cells. `AGG_ERODE`, `REINST_PREM`, `ON_LEVEL`, and `RATE_INDEX` are included on this basis.
-   **Does it compress a common formula?** Standardizing a name to a repeated expression across multiple workbooks and models can 1) allow for some efficiency, and 2) remove a source of error (e.g. the formula being mistyped). For example, `XOL_RECOV` replaces `=MAX(0, MIN(loss-ret, lim))`; and `XINTERP` allows for on-the-fly interpolation.

------------------------------------------------------------------------

## Install in 60 Seconds

1.  Download **`GILL_V0.1.bas`** from this repository.
2.  Open your target workbook and press <kbd>Alt</kbd>+<kbd>F11</kbd> to open the VBA editor.
3.  Select **File** $\rightarrow$ Import File... and choose `GILL_V0.1.bas`.
4.  Close the VBA Editor. Press <kbd>Alt</kbd>+<kbd>F8</kbd> $\rightarrow$ select `GILL_Install` $\rightarrow$ Run).

You should see confirmation: *"General Insurance Lambda Library v0.1 installed."* Start typing `=XOL_` in any cell to trigger autocomplete. `=XOL_RECOV` should pop up.

You should also see a sheet named 'GILL v0.1', along with a list of function definitions and why #N/A appears whenever it does in the output. Probably a good idea to keep a readable copy for peer review.

*Very cool,* but does it run? Check this using **GILL_SelfTest()** -- all tests should pass (more on this below). The self-test checks the built-in versions, so if you've edited a function or installed your own from a file, some checks may fail (which is to be anticipated).



### **Requirements:**

-   **Microsoft 365 or Excel 2024+:** (utilizes array functions like `MAP`, `SCAN`, `REDUCE`, `MAKEARRAY`, `SORTBY`, `TAKE`, `DROP`, `VSTACK`, `TOCOL`).
-   **Desktop Installation of Excel:** Excel on the web doesn't support VBA, so individual functions would have to be defined manually. So currently, installation must happen in desktop Excel. Once installed, however, the workbook calculates seamlessly in Excel Online. Note that this did not appear to be compatible with Google Sheets.

### Optional: Keep the installer one click away

1.  Open a blank workbook, import `GILL_V0.1.bas` as above.
2.  **File** $\rightarrow$ Save As $\rightarrow$ Excel Add-In (`.xlam`).
3.  **File** $\rightarrow$ Options $\rightarrow$ Add-ins $\rightarrow$ Go..., tick your new add-in.
4.  Add `GILL_Install` to the Quick Access Toolbar (**Options** $\rightarrow$ Quick Access Toolbar $\rightarrow$ Choose commands from: Macros).

Unlike Solver, this doesn't make the functions available in every workbook at once. Excel ties these functions to a workbook, so the add-in gives you a button, and the functions go into whichever workbook you click it in.

------------------------------------------------------------------------

## Macros in GILL v0.1

| Macro | What it does |
|------------------------------------|------------------------------------|
| `GILL_Install` | Adds or refreshes the 14 built-in functions in the active workbook. |
| `GILL_InstallFromFile` | Adds or refreshes all 14 functions in the active workbook based on a custom list from a text file that you pick. |
| `GILL_SelfTest` | Evaluates every function against known answers (about 3 dozen checks). |
| `GILL_Export` | Writes the **current** definitions to `GILL_export.txt`, in the format `GILL_InstallFromFile` reads back (so it's a nice round trip). |
| `GILL_Remove` | Deletes only names created by this module (tracked via Name Manager *Comment* tags), as well as the info sheet. |
| `GILL_Help` | Opens documentation, i.e. this page. |

Note 1: There are helper macros and functions powering these, but only these 6 show up in the Alt+F8 list, since they don't have any parameters. Excel hides anything that has even an optional parameter.
Note 2: If you have a bad custom functions file, GILL_InstallFromFile will tell you where things broke.

------------------------------------------------------------------------

## Function Reference

Optional arguments are shown in `[brackets]`, consistent with Excel standards.

### 0. General

| Function | What does it DO? | Syntax | Example & Result |
|------------------|------------------|------------------|------------------|
| **XINTERP** | Interpolates a value from known data points using linear, log-linear, or nearest-neighbor methods | `XINTERP(x, x_known, y_known, [method])` | `XINTERP(2.5,{1;2;3;4},{10;20;30;40})` $\rightarrow$ **25** |


### 1. Pricing

| Function | What does it do? | Syntax | Example & Result |
|------------------|------------------|------------------|------------------|
| **LAYER_LOSS** | Calculates expected layer loss using ILFs | `LAYER_LOSS(el, attach, layer_lim, ilf_lims, ilf_vals, [method])` | $\$5\text{M xs }\$1\text{M}$ on a standard ILF table $\rightarrow$ **307,000** |
| **MBBEFD** | Returns the Swiss Re / Bernegger exposure curve value G(d) | `MBBEFD(d, b, g)` | `MBBEFD(0.5, 2, 3)` $\rightarrow$ **0.6262** |
| **CRED_Z** | Calculates Limited Fluctuation or Buhlmann credibility Z factor | `CRED_Z(n, [method], [k])` | `CRED_Z(300)` $\rightarrow$ **0.5266** |
| **ON_LEVEL** | Applies on-level factors using historical rate change arrays | `ON_LEVEL(pol_eff, rate_dates, rate_chgs, [as_of])` | $+5\%$ then $+8\%$ rate changes $\rightarrow$ **1.134** |
| **RATE_INDEX** | Compiles cumulative rate level index from a history of rate changes | `RATE_INDEX(changes, [base])` | `RATE_INDEX({0.05;0.08})` $\rightarrow$ `{1.05; 1.134}` |

### 2. Reserving

| Function | What does it do? | Syntax | Example & Result |
|------------------|------------------|------------------|------------------|
| **TRIANGLE** | Converts tabular loss data (origin, dev, value) into a 2D triangle array | `TRIANGLE(origin, dev, values, [cum])` | Flat columns $\rightarrow$ 2D triangle (`cum=TRUE` for cumulative). |
| **LDF_SELECT** | Selects LDFs using simple, medial, or volume-weighted averages | `LDF_SELECT(factors, [method], [n_periods], [weights])` | `LDF_SELECT({1.892;1.756;2.104;1.834;1.901},"medial")` $\rightarrow$ **1.8757** |
| **ULTIMATE** | Consolidates CL, BF, and Cape Cod reserving methods | `ULTIMATE(reported, cdf, [method], [elr], [premium])` | `ULTIMATE(8500000,1.25,"bf",0.65,15000000)` $\rightarrow$ **10,450,000** |
| **ALAE_LOAD** | Applies ALAE loading to indemnity losses (pro_rata, flat, capped) | `ALAE_LOAD(indem, method, param, [cap])` | `ALAE_LOAD(700000,"capped",0.15,100000)` $\rightarrow$ **100,000** |

*Methods:* `LDF_SELECT` supports `simple`, `medial`, and `volume`. `ULTIMATE` supports `cl`, `bf`, and `cc`. 

### 3. Reinsurance

| Function | What does it DO? | Syntax | Example & Result |
|------------------|------------------|------------------|------------------|
| **XOL_RECOV** | Calculates per-occurrence excess of loss recovery | `XOL_RECOV(losses, retention, limit, [loss_cap])` | `XOL_RECOV(8500000,5000000,5000000)` $\rightarrow$ **3,500,000** |
| **AGG_ERODE** | Tracks aggregate limit/deductible erosion across a sequence of losses | `AGG_ERODE(losses, retention, limit, aggregate, [agg_type])` | $\$2\text{M xs }\$3\text{M}$ under a $\$5\text{M}$ AAD $\rightarrow$ `{0;0;0;2M;2M}` |
| **REINST_PREM** | Calculates sequential reinstatement premium (pro rata or flat) | `REINST_PREM(losses, ret, lim, base_prem, n_reinst, [rate], [type])` | $\$5\text{M xs }\$5\text{M}$, 2 reinstatements $\rightarrow$ **1,000,000** |
| **CORRIDOR** | Blends standard and inside-corridor cession percentages | `CORRIDOR(lr, attach, width, base_cess, [corr_cess])` | `CORRIDOR(0.75,0.7,0.1,0.5,0)` $\rightarrow$ **0.4667** |

## Design Choices:

1. *XINTERP:* Doesn't extrapolate - throws an #N/A if it goes out of bounds. If necessary (e.g. flat extrapolation), it's best to clamp with MIN/MAX outputs for full transparency.
2. *CORRIDOR:* Uses the weighted-average cession rate across the entire loss ratio (as opposed to incremental dollars of loss that fall inside the corridor).

**Editing these functions:** At the end of the day, these are defined names in the Name Manager. You can open that or the Advanced Formula Environment up and try changing these yourself (watch out for the bracket-matching!), and you can keep using your edited version until you re-run GILL_Install. If you want to save your changes for posterity, GILL_Export is always an option.

------------------------------------------------------------------------

## Robustness

To prevent silent failures in data pipelines, GILL tries to fail visibly instead of returning a wrong number:

| Situation | Behavior |
|------------------------------------|------------------------------------|
| Blank cell | Read as `0` (except in `XINTERP`, `LAYER_LOSS`, `LDF_SELECT`) |
| Text in a function that works *cell by cell* in an array | Returns `#N/A` for **that cell** only |
| Text in a function that keeps a *running total* (`AGG_ERODE`, `RATE_INDEX`, `ON_LEVEL`, `REINST_PREM`) | Returns `#N/A` for the **whole result**, because one unknown value makes every later total unknown |
| Missing required argument or unknown method name | Returns `#N/A` |

An initial thought was to write custom errors for easier debugging, but any successive functions wrapping these results in ISNA()/IFNA() might benefit from this more recognizable output (just #N/A). If it might be beneficial for users to see the specific failure modes instead, I could write text descriptions to deal with this.

### Self-Tests

Right now, 35 tests have been defined in the following categories to ensure silent failures are minimized:

1. XINTERP: Interpolation, unsorted tables, blank table rows, and `#N/A` below the table.
2. RATE_INDEX, ON_LEVEL: Non-numeric entries in running index, an unsorted rate history, and text in the rate history.
3. CRED_Z: Limited fluctuation Z
4. LAYER_LOSS, TRIANGLE: Properly handled blank table rows in ILF schedules and accurate boundary masking across origin/development dimensions.
5. MBBEFD: Verified against general Bernegger curves and boundary conditions (e.g. G(1) = 1).
6. ALAE_LOAD: Verified to preserve proper expense loading for 0-indemnity (closed, no-pay) claims under flat method.
7. LDF_SELECT: Tested for automatic fallback to simple averages at n<=2 for medial selections, rejection for missing volume weights, skipping empty factor cells.
8. ULTIMATE: Chain-ladder and BF results, #N/A for a missing ELR or premium, unrecognized method, text in reported losses - goal is that there's no silent fallback to another calculation.
9. AGG_ERODE, REINST_PREM: Tested to check dirty/text inputs propagate #N/A across array, i.e. text is not silently coerced to 0.
10. CORRIDOR, XOL_RECOV: Verified for correct proportional loss cession blending and per-occurrence excess of loss recovery calculations.

### Error reporting

While the functions aim to be robust within reason, there may be unforeseen cases that produce #DIV, #N/A or other errors. Of course, the more insidious kind of error would leave no trace. Feel free to report these in discussion; the more context you could provide around the error, the more helpful it would be to diagnose the issue.

### Other Wrenches Caught So Far

1. (v0.1) Lambda line lengths leaned long, so they got broken into multi-line definitions via VBA's '_' syntax.
2. (v0.1) Blanks are skipped within custom functions.



------------------------------------------------------------------------

## What Next?

1.  I'm still deciding on adding 'fit.dist' functionality (like fitdistrplus in R), but code it into a button on a ribbon rather than within a lambda.
2.  Adding a TypeScript or JavaScript version, so this vibes with Excel Web as well.
3.  If there is merit to splitting this out into a new tab, a more full-fledged .xlam that has some UI (say, making the Import/Export functionality permanent within your workbook) could be something, though this could make things bulkier.

## Contributing & License

Contributions, corrections, and high-frequency formula additions are welcome. Please ensure new submissions meet the design bars outlined above.

Licensed under the GNU General Public License v3.0 (GPLv3). See LICENSE for details.
