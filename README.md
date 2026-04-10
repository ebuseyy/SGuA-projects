# SGuA-projects

## Repository for the ORC / DORC Experimental Codes and Revision Materials

This repository contains the source files, notebooks, and archived implementation materials associated with the study on optimization-based overcurrent relay coordination (ORC) and directional overcurrent relay coordination (DORC), including both **pre-revision** and **post-revision** stages of the work.

The repository was organized to improve transparency, reproducibility, and traceability throughout the revision process.

---

## Repository Structure

```text
.
├── AfterRevision/
│   ├── ORC_ABC_30runs.ipynb
│   ├── ORC_BOA_30runs.ipynb
│   ├── ORC_GA_30runs.ipynb
│   ├── ORC_GWO_30runs.ipynb
│   ├── ORC_LCA_30runs.ipynb
│   ├── ORC_WOA_30runs.ipynb
│   ├── SGuA-ISGuA-CMA_ES-L_SHADE.ipynb
│   ├── dorc_microgrid_9bus_paper_anchored_dynamic_init.ipynb
│   ├── orc_meta_common.py
│   └── README.md
├── BeforeRevision/
│   ├── ABC.ipynb
│   ├── BOA.ipynb
│   ├── GWO.ipynb
│   ├── OzgeSGuA_CEC_InitPop.ipynb
│   ├── FGA.cs
│   ├── FGA.Designer.cs
│   ├── FGA.resx
│   ├── Form1.cs
│   ├── Form1.Designer.cs
│   ├── Form1.resx
│   └── README.md
└── README.md
```

---

## Purpose of This Repository

This repository was prepared to document the computational materials used during different phases of the study. It includes:

- archived pre-revision implementations
- revised post-revision notebooks
- algorithm-specific ORC experiments
- comparative experiments involving multiple metaheuristic methods
- a revised DORC microgrid case study
- supporting source files used during the manuscript revision process

The repository is intended primarily as a **research archive and reproducibility companion** to the associated manuscript.

---

## Main Folder Descriptions

### 1. `BeforeRevision/`

This folder contains earlier or archived materials from the **pre-revision stage** of the study.

It includes:

- early Python notebooks for selected ORC experiments
- original C# implementations used in the earlier development stage
- archived benchmark-related materials

Important notes:

- `FGA.cs` corresponds to the original **SGuA** implementation developed in C#
- `Form1.cs` corresponds to the original **LCA** implementation developed in C#
- these files were retained mainly for archival and historical reference
- file naming reflects the original development environment and was preserved to maintain continuity with the author’s earlier workflow

For more details, see: `BeforeRevision/README.md`

### 2. `AfterRevision/`

This folder contains the **post-revision** materials prepared after reviewer comments and manuscript updates.

It includes:

- revised ORC notebooks for ABC, BOA, GA, GWO, LCA, and WOA
- the updated comparison notebook involving SGuA, ISGuA, CMA-ES, and L-SHADE
- the revised DORC microgrid 9-bus case study
- a shared Python helper module used across multiple ORC notebooks

This folder should be considered the **main reference point** for the revised stage of the study.

For more details, see: `AfterRevision/README.md`

---

## How to Interpret the Repository

The repository is intentionally divided into two major stages:

| Folder | Meaning |
|---|---|
| `BeforeRevision/` | Archived materials from the earlier stage of the study |
| `AfterRevision/` | Revised and reorganized materials prepared after reviewer comments |

This separation was made to preserve the development history and to distinguish earlier implementations from the updated experimental structure used in the revised manuscript.

---

## Manuscript-Relevant Content

In general:

- the **main revised ORC experiments** are located in `AfterRevision/`
- the **main revised comparison framework** is provided in `AfterRevision/SGuA-ISGuA-CMA_ES-L_SHADE.ipynb`
- the **revised DORC case study** is provided in `AfterRevision/dorc_microgrid_9bus_paper_anchored_dynamic_init.ipynb`
- the **earlier archived implementations** are preserved in `BeforeRevision/`

If manuscript table or figure numbering changes during revision, the repository should be interpreted based on experiment groups and file names rather than fixed numbering alone.

---

## Reproducibility Notes

To reproduce the Python-based experiments:

1. clone or download the repository
2. keep the folder structure unchanged
3. install the required Python packages
4. open the notebooks in Jupyter Notebook or JupyterLab
5. execute cells sequentially from top to bottom

Some notebooks in `AfterRevision/` rely on:

- `orc_meta_common.py`

Therefore, this file should remain in the same folder as the notebooks unless the import paths are modified.

---

## Software and Execution Environment

This repository contains a mixture of:

- **Python / Jupyter Notebook** based experimental files
- **C# Windows Forms** based legacy files from the earlier stage

Accordingly:

- Python notebooks should be executed in a Jupyter environment
- C# files should be opened in a compatible Visual Studio environment if needed

The C# files are included mainly for archival transparency and historical continuity.

---

## Scope and Limitations

This repository reflects the author’s actual experimental workflow and revision history. Therefore:

- file names may follow the original working structure
- some files are retained in their legacy form
- the repository is not presented as a polished software package
- instead, it is presented as a structured research archive accompanying the manuscript

---

## Recommended Reading Order

For readers interested mainly in the revised study, the following order is recommended:

1. read this main `README.md`
2. open `AfterRevision/README.md`
3. inspect the ORC notebooks in `AfterRevision/`
4. inspect `SGuA-ISGuA-CMA_ES-L_SHADE.ipynb`
5. inspect the DORC notebook
6. use `BeforeRevision/` only if archival comparison or development history is needed
