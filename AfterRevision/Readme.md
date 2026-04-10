# AfterRevision

This folder contains the revised implementations and experimental notebooks prepared after the manuscript revision process. It was organized to support transparency, reproducibility, and clearer correspondence between the revised manuscript and the associated computational experiments.

## Purpose of This Folder

The files in this directory represent the **post-revision stage** of the study. They include:

- revised ORC experiments for multiple comparison algorithms,
- the updated comparative framework involving SGuA, ISGuA, CMA-ES, and L-SHADE,
- the revised DORC microgrid case study based on the 9-bus system,
- a shared Python utility module used across multiple notebooks.

This folder should be interpreted together with the revised manuscript. Earlier or archived materials belong to the separate `BeforeRevision` folder.

---

## File Descriptions

### 1. ORC notebooks (30 independent runs)

These notebooks contain revised ORC experiments for individual comparison algorithms. Each notebook was prepared to perform repeated runs and report the corresponding relay setting results and fitness values.

- `ORC_ABC_30runs.ipynb`  
  Artificial Bee Colony (ABC) based ORC experiments.

- `ORC_BOA_30runs.ipynb`  
  Butterfly Optimization Algorithm (BOA) based ORC experiments.

- `ORC_GA_30runs.ipynb`  
  Genetic Algorithm (GA) based ORC experiments.

- `ORC_GWO_30runs.ipynb`  
  Grey Wolf Optimizer (GWO) based ORC experiments.

- `ORC_LCA_30runs.ipynb`  
  League Championship Algorithm (LCA) based ORC experiments.

- `ORC_WOA_30runs.ipynb`  
  Whale Optimization Algorithm (WOA) based ORC experiments.

### 2. Main revised comparison notebook

- `SGuA-ISGuA-CMA_ES-L_SHADE.ipynb`  
  This notebook contains the main revised comparison framework involving SGuA, ISGuA, CMA-ES, and L-SHADE. It corresponds to the central post-revision analysis where the proposed and reference methods were compared in a unified manner.

### 3. Revised DORC case study notebook

- `dorc_microgrid_9bus_paper_anchored_dynamic_init.ipynb`  
  This notebook contains the revised directional overcurrent relay coordination (DORC) case study for the 9-bus microgrid system.

### 4. Shared utility module

- `orc_meta_common.py`  
  This module includes helper functions, shared routines, and common definitions used by the ORC notebooks. It should remain in the same directory as the notebooks if they are executed directly.

---

## Correspondence Between Manuscript Content and Repository Files

The following mapping is provided to make it easier to identify which file supports which part of the revised manuscript.

| Manuscript content | Related file |
|---|---|
| Revised ORC results for ABC | `ORC_ABC_30runs.ipynb` |
| Revised ORC results for BOA | `ORC_BOA_30runs.ipynb` |
| Revised ORC results for GA | `ORC_GA_30runs.ipynb` |
| Revised ORC results for GWO | `ORC_GWO_30runs.ipynb` |
| Revised ORC results for LCA | `ORC_LCA_30runs.ipynb` |
| Revised ORC results for WOA | `ORC_WOA_30runs.ipynb` |
| Revised comparison involving SGuA, ISGuA, CMA-ES, and L-SHADE | `SGuA-ISGuA-CMA_ES-L_SHADE.ipynb` |
| Revised DORC microgrid 9-bus study | `dorc_microgrid_9bus_paper_anchored_dynamic_init.ipynb` |
| Shared ORC helper routines | `orc_meta_common.py` |

If table numbers, figure numbers, or section numbers are updated during manuscript revision, this mapping may be interpreted at the level of experiment group rather than fixed numbering.

---

## Suggested Execution Order

For practical use, the following order is recommended:

1. Keep `orc_meta_common.py` in the same folder.
2. Run the individual ORC notebooks as needed:
   - `ORC_ABC_30runs.ipynb`
   - `ORC_BOA_30runs.ipynb`
   - `ORC_GA_30runs.ipynb`
   - `ORC_GWO_30runs.ipynb`
   - `ORC_LCA_30runs.ipynb`
   - `ORC_WOA_30runs.ipynb`
3. Run the main revised comparison notebook:
   - `SGuA-ISGuA-CMA_ES-L_SHADE.ipynb`
4. Run the DORC case study notebook:
   - `dorc_microgrid_9bus_paper_anchored_dynamic_init.ipynb`

---

## Reproducibility Notes

To reproduce the experiments:

1. place all files in the same directory structure,
2. install the required Python packages,
3. open the notebooks in Jupyter Notebook or JupyterLab,
4. execute the cells sequentially from top to bottom.

Because some notebooks use shared functions, moving files to different directories without updating import paths may cause execution errors.

---

## Folder Scope and Interpretation

This folder contains the **revised and reorganized experimental materials** prepared after reviewer comments and manuscript revision. It is intended to document the post-revision computational workflow more clearly than the earlier archived materials.

The `BeforeRevision` folder separately contains earlier or legacy implementations retained for archival and historical reference.

---

## Important Note

The files in this folder were kept close to the author’s experimental workflow and naming history. For this reason, file names reflect the revision process and experimental grouping rather than a completely standardized software package structure.
