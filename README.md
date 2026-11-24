# Circularity rarely speaks — replication package

Code, coding scheme and derived data for the manuscript **"Circularity rarely speaks: R-strategies in visitor reviews across ten emerging economies"**, submitted to *Production* (ABEPRO), Thematic Section linked to IJCIEOM 2027.

The study measures how often visitors' online reviews of tourism suppliers report circular economy practices (R-strategies), which strategies they report, and how those reports relate to ratings. Country estimates are pooled with a single-corpus random-effects meta-analysis.

## Main results (reproduced by the notebook)

| Estimate | Value |
|---|---|
| Prevalence of circular reports (10 countries) | 0.73% (95% CI 0.45–1.18%), I² = 93% |
| Without China (k = 9) | 0.68% (0.40–1.15%) |
| By city (k = 29) | 0.69% (0.52–0.93%) |
| OR circular report, creative vs mass-market | 4.23 (2.51–7.14) |
| OR 5-star rating given a circular report | 0.58 (0.43–0.78) |
| OR 1–2 star rating given a repair report | 3.20 (1.94–5.27) |
| OR 5-star rating given a local/artisanal report | 1.53 (0.99–2.35), p = 0.054 |
| Inter-coder agreement (n = 296) | κ = 0.86 (relevance), κ = 0.90 (strategy) |

## Repository structure

```
notebooks/
  circularity_meta_analysis.ipynb   full pipeline, executed, with outputs
scripts/
  build_coding_sheet.py             builds the double-coding workbook from the raw corpus
data/
  lexicon.csv                       seed lexicon (regular expressions) by R-strategy group, in priority order
  noise_filters.csv                 idiomatic uses removed before coding
  codebook.md                       coding rules (relevance, strategy, actor, valence)
  coded_passages_coder1.csv         codes for all 1,176 passages (no review text)
  coded_passages_coder2.csv         codes for the 296-passage double-coded subsample (no review text)
results/
  resultados_meta_analise.xlsx      pooled estimates, Table 2, Table 3, country estimates
  Figure1.png                       forest plot (300 dpi)
references/
  references.md                     works cited in the manuscript, with DOIs
```

## How to reproduce

1. Install the requirements: `pip install -r requirements.txt`.
2. Place the review corpus at `data/reviews_raw.csv` (see *Data availability*).
3. Open `notebooks/circularity_meta_analysis.ipynb` and run all cells. Results are written to `results/`.

The notebook first re-derives the coding set from the corpus with the fixed seed (20261119) and checks that it matches the identifiers in `data/coded_passages_coder1.csv`. Everything downstream (kappa, indicators, sampling weights, meta-analysis, tables and figure) is computed from the coded files.

## Method in brief

- **Corpus:** 54,687 Google Maps reviews of 11,509 tourism suppliers in 29 cities of ten emerging economies; 51,128 with text. Texts are served in Portuguese (original or platform translation).
- **Retrieval:** word-boundary regular expressions grouped as R0–R2, R3, R4, R5–R6, R7, R8, plus three comparison groups (local supply chain, artisanal production, generic sustainability claim). Idiomatic uses are filtered out.
- **Sampling:** census of all candidate passages except artisanal-only matches, of which 200 of 629 were sampled at random, proportionally by segment, and reweighted by the inverse sampling fraction.
- **Coding:** relevance, main and secondary strategy, actor and valence (`data/codebook.md`); a 25% stratified subsample was double-coded.
- **Meta-analysis:** countries as replications; logit prevalences and log odds ratios; DerSimonian–Laird τ²; Hartung–Knapp confidence intervals; 95% prediction intervals.

## Data availability

`reviews_raw.csv` contains review text and reviewer names from Google Maps. Platform terms do not allow redistribution, so the file is not included and is listed in `.gitignore`. The coded files in `data/` contain only identifiers, metadata (country, city, segment, rating), the matched seed terms and the codes. They do not contain review text or reviewer names.

## Related work on the same corpus

The corpus was collected under a sampling frame established in a prior study of the same supplier population. This study builds its own lexicon, coding scheme and variables; no measure from earlier work is reused.

## License

Code: MIT License. Derived data and documentation: CC BY 4.0.
