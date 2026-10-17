# Spatial Clustering of Flood Impact in Thailand (2019–2024)

Group Assignment #2 – Analytics on Complex Data (Data Mining)

## Run

```bash
python3 -m venv .venv && .venv/bin/pip install pandas numpy scikit-learn matplotlib seaborn geopandas \
    libpysal esda mapclassify openpyxl lxml html5lib jupyter jupytext
.venv/bin/python src/build_dataset.py                 # raw files -> data/processed/
.venv/bin/jupytext --to ipynb flood_analysis.py -o flood_analysis.ipynb
.venv/bin/jupyter nbconvert --to notebook --execute --inplace flood_analysis.ipynb
```

`flood_analysis.py` (jupytext percent format) is the source of the notebook; edit it, then regenerate the `.ipynb`.

Slides: `cd slides && npm install && NODE_PATH=$PWD/node_modules PPTX_SKILL=<pptx skill dir> node build_deck.js`
(numbers come from `data/processed/results.json`, figures from `figures/`).

## Data (data/raw/)

| Folder | Source | Content |
|---|---|---|
| `ddpm/` | DDPM, catalog.disaster.go.th dataset `dpm-gd027` | yearly flood statistics 2562–2567 (6 different schemas) |
| `rain/` | HII / ThaiWater, data.go.th `spatial-rain` | monthly areal rainfall per province, 30-yr normal |
| `dopa/` | DOPA, data.go.th `dopa-star` | registered population per province (HTML saved as .xls) |
| `ldd/` | Land Development Dept., data.go.th `ldd_21_04` | recurrent-flood area per sub-district (validation reference) |
| `geo/` | github.com/chingchai/OpenGISData-Thailand | 77 province polygons |

The GISTDA recurrent-flood API needs a registered API key, so the LDD recurrent-flood statistics are used as the reference map.

## Outputs

- `data/processed/province_year.csv` – 77 provinces × 6 years, harmonised
- `data/processed/province_results.csv` – features, clusters, LISA class per province
- `data/processed/results.json` – every number quoted in the slides
- `figures/` – all charts and maps
- `slides/Flood_Clustering_Thailand.pptx` – presentation (Thai speaker notes for the video)
- `report/Flood_Clustering_Report.docx` / `.pdf` – written report (`NODE_PATH=$PWD/slides/node_modules node report/build_report.js`, then update the table of contents in Word)
