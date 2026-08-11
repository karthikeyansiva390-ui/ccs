# CCS Screening & Investment Decision Framework — Robust Input Normalization Version

## What changed in this version

The main weakness in the earlier implementation was that it expected the raw field-cell content to already be directly comparable with the reference-cell content.

This version adds a dedicated **Input Normalization and Reference Canonicalization Layer** before screening.

### Numerical/measurable inputs

The engine now:

1. Reads the raw Excel/CSV value.
2. Detects blank/N/A/-/unknown-style missing values.
3. Strictly validates whether the content is numeric.
4. Detects a unit embedded in the cell, if present.
5. Uses the Field `Unit/Type` as the fallback unit.
6. Reads the Reference `Data Type` as the target unit.
7. Converts compatible units using `pint`.
8. Produces one canonical numeric value.
9. Tests that canonical value against the reference hard cut-off.
10. Searches ALL five SAW numerical ranges.
11. Assigns the matching SAW score only when a range is actually found.

Example:

`1,500 mD` → numeric value `1500` → reference unit `mD` → canonical value `1500 mD` → hard-cutoff check → SAW range lookup.

If a value is numerically valid but does not fall in any SAW range, it is **not assigned an artificial score**.

### Qualitative inputs

The engine now reads the qualitative vocabulary from:

- Hard Cut-Off Values
- SAW Score 1 text
- SAW Score 2 text
- SAW Score 3 text
- SAW Score 4 text
- SAW Score 5 text

The field value is normalized before comparison:

- case-insensitive
- Unicode-safe
- extra spaces ignored
- punctuation/spacing differences normalized
- harmless hyphen/slash differences normalized

Example:

`HIGHLY SUITABLE`

and

`Highly-Suitable`

can resolve to the same canonical reference phrase.

The engine then searches the **complete SAW text vocabulary**, rather than only searching the hard-cutoff text.

### Fuzzy matching

The package `rapidfuzz` is used as a conservative final fallback.

It does NOT blindly convert every sentence into a reference value.

- normalized exact match → accepted
- highly similar match (90%+) → accepted
- lower-confidence match → rejected and reported as unmatched

This is deliberate. There is no honest way for software to guarantee "100% semantic equivalence" between two arbitrary engineering sentences. A false positive is more dangerous to a screening framework than a field being flagged for manual correction.

### Missing / inappropriate values

The engine treats these as no usable input:

- blank
- `N/A`
- `NA`
- `N.A.`
- `-`
- `--`
- `—`
- `Not Available`
- `Not Applicable`
- `Unknown`
- `None`
- `Null`
- `Nil`
- `Not Specified`
- `Not Reported`

For numerical fields, arbitrary text such as `unknown 25 mD` is rejected rather than incorrectly extracting `25`.

## Phase 1 logic

```text
Field Data
   ↓
Input normalization / canonicalization
   ↓
Missing or invalid?
   ├── YES → FAIL
   └── NO
        ↓
Hard Cut-Off
        ↓
Fail?
   ├── YES → eliminate + reason
   └── NO
        ↓
SAW classification
        ↓
No SAW range/text match?
   ├── YES → FAIL + reason
   └── NO
        ↓
SAW score × sub-parameter AHP weight
        ↓
Sum within each common parameter
        ↓
Common parameter score × overall AHP weight
        ↓
Final Phase-1 SAW score
        ↓
Rank fields
```

## Important qualitative rule

For qualitative parameters, the field input is first mapped to the reference vocabulary.

The hard-cutoff test is still performed before the final SAW score is accepted.

If the hard-cutoff cell itself is an explicit qualitative threshold and it can also be mapped to one of the SAW categories, the implementation can compare the field's SAW category against that threshold. This is useful when a reference sheet expresses a qualitative threshold such as "Suitable" while the SAW categories contain "Moderately Suitable", "Suitable", "Highly Suitable", etc.

Do not rely on this threshold interpretation unless it matches your thesis methodology. If a parameter is strictly categorical rather than ordinal, use explicit accepted hard-cutoff text values.

## Phase 2

The existing Phase-2 implementation remains:

1. Expected vs Actual CAPEX/OPEX gate.
2. Economic analysis:
   - NPV
   - IRR
   - Payback Period
3. Sensitivity analysis:
   - CAPEX
   - OPEX
   - Discount Rate
   - Inflation Rate
   - CO2 Injection Rate
   - Project Lifetime
   - Carbon Credits

Phase-2 gate-eliminated fields can still be manually overridden into economic analysis, as required by the framework.

## New diagnostic table

After Phase 1 runs, expand:

**View value/text normalization and reference matching diagnostics**

This shows:

- raw field value
- field unit
- normalized/canonical value
- normalization note
- hard-cutoff value
- selected SAW score
- matched SAW reference text/range
- SAW match similarity
- pass/fail

This is extremely useful for debugging your reference and field Excel files before trusting the ranking.

## Input workbook format

### Reference workbook

Supported sheets:

- Overall
- Technical
- Environmental
- Regulatory
- Long Term Operation
- Risk

Category sheet:

- Column A: Parameter
- Column B: Data Type
- Column C: Hard Cut Off Values
- Column D: AHP Weightage (%)
- E:F: SAW 1 start/end
- G:H: SAW 2 start/end
- I:J: SAW 3 start/end
- K:L: SAW 4 start/end
- M:N: SAW 5 start/end

### Field workbook

Supported sheets:

- Technical
- Environmental
- Regulatory
- Long Term Operation
- Risk
- Economics

Category sheets:

- Parameter
- Value
- Unit/Type

Economics:

- Actual CAPEX
- Actual OPEX

## Installation

```bash
pip install -r requirements.txt
streamlit run app.py
```

## GitHub / Streamlit Community Cloud

Upload:

- `app.py`
- `ccs_engine.py`
- `requirements.txt`
- `README.md`
- `.gitignore`

Deploy `app.py` as the main Streamlit file.

## Critical testing recommendation

Before using the framework for dissertation results, create a small controlled reference workbook with known values and intentionally test:

1. exact numerical value inside each SAW range;
2. numerical value outside every SAW range;
3. numerical value with commas;
4. numerical value with units inside the cell;
5. unit conversion;
6. blank field;
7. `N/A`;
8. `-`;
9. qualitative exact match;
10. qualitative different capitalization;
11. qualitative punctuation variation;
12. qualitative text matching SAW 1–5;
13. qualitative text below the fuzzy threshold;
14. a field passing the hard gate but having no SAW match.

The new diagnostics table should be used to verify every case.

## Academic caution

The normalization package is an input interpretation layer; it must not be described as making arbitrary field data "100% correct."

The strongest defensible statement is:

> "The framework converts heterogeneous field inputs into a standardized representation using deterministic normalization, unit conversion and conservative reference-vocabulary matching before applying the screening rules."

That is technically much stronger than claiming that any arbitrary sentence can be guaranteed to have exactly the same meaning as a reference sentence.


## Final Phase-1 input logic

The application uses only the reference datasheet and manually entered field data.
No fuzzy matching, semantic matching, automatic unit conversion, or field-datasheet
normalization package is used.

For each numerical/measurable sub-parameter:
1. The entered numerical value is checked against the five SAW ranges first.
2. If it falls in a SAW range, that SAW score is used immediately and multiplied by the sub-parameter AHP weight.
3. If it does not fall in any SAW range, the hard-cutoff criterion is checked.
4. If the hard cut-off fails, the field is eliminated.
5. If the hard cut-off passes but there is still no SAW score, the engine does not invent a score; it reports that the parameter is unscorable.

For qualitative sub-parameters:
- The user selects a reference phrase/sentence from a dropdown.
- The dropdown is populated from the reference SAW text values and hard-cutoff text.
- SAW text is checked first; hard-cutoff text is checked only when no SAW text matches.
- Capitalization and harmless punctuation/spacing differences are ignored.

The manual field-entry dialog has exactly six tabs:
Technical, Environmental, Regulatory, Long Term Operation, Risk, and Economics.
Actual CAPEX and Actual OPEX are entered in the Economics tab, not at the bottom
of the other parameter tabs.


## Hard-cutoff-only PASS — final rule

A Phase-1 sub-parameter has three possible outcomes:

1. **SAW match:** the input coincides with a reference SAW range/text -> PASS and receive that exact SAW score (1–5).
2. **Hard-cutoff-only match:** the input does not coincide with any SAW range/text, but satisfies the reference hard cut-off -> PASS, but **NO SAW SCORE** is assigned.
3. **Failure:** the input does not coincide with any SAW range/text and also does not satisfy the hard cut-off -> the field is eliminated.

The software never assigns an artificial SAW score of 0, 1, 2, 3, 4 or 5 when the input has no matching SAW reference range/text.

## Final regression-tested build

This build has been regression-tested against the supplied reference workbook.
The tests cover:
- reference XLSX loading and all five CCS parameter groups;
- Unicode comparison operators such as `≥` and `≤`;
- scientific notation such as `1*10^-6` and `5×10⁻⁷`;
- open-ended SAW ranges such as `>100` and `<1*10^-9`;
- SAW-first scoring;
- hard-cutoff-only PASS with **no SAW score**;
- hard-cutoff failure/elimination;
- full Phase-1 score aggregation without adding `None` values;
- Phase-2 CAPEX/OPEX gate and economic cash-flow calculation.

Run locally with:

```bash
python test_ccs_framework.py
```

The application UI displays the Phase-1 manual-entry Economics tab as:
- Actual CAPEX — **USD**
- Actual OPEX — **USD/Year**
