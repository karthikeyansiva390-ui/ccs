from pathlib import Path
import importlib.util

ENGINE = Path(__file__).with_name('ccs_engine.py')
REFERENCE = Path('/mnt/data/Reference Data sheet (with values)(2).xlsx')

spec = importlib.util.spec_from_file_location('ccs_engine', ENGINE)
engine = importlib.util.module_from_spec(spec)
spec.loader.exec_module(engine)

reference = engine.parse_reference_xlsx(REFERENCE.read_bytes())

# 1. Verify the actual reference workbook loads completely.
assert set(reference['categories']) == set(engine.CATEGORIES)
assert len(reference['categories']['Technical']) == 32
assert len(reference['categories']['Environmental']) == 8
assert len(reference['categories']['Regulatory']) == 9
assert len(reference['categories']['Long Term Operation']) == 10
assert len(reference['categories']['Risk']) == 9

# 2. Verify scientific notation and Unicode comparison operators from the workbook.
assert abs(engine.parse_number_strict('1*10^-6') - 1e-6) < 1e-15
assert abs(engine.parse_number_strict('5×10⁻⁷') - 5e-7) < 1e-15
assert engine.parse_operator_cutoff('≥20') == ('>=', 20.0)
assert engine.parse_operator_cutoff('<1*10⁻⁹')[0] == '<'
assert abs(engine.parse_operator_cutoff('<1*10⁻⁹')[1] - 1e-9) < 1e-15

# 3. Verify open-ended SAW ranges actually work.
caprock = next(x for x in reference['categories']['Technical'] if x['parameter'] == 'Caprock permeability')
r = engine.evaluate_parameter(2e-10, 'mD', caprock)
assert r['pass'] and r['saw_score'] == 5, r

thickness = next(x for x in reference['categories']['Technical'] if x['parameter'] == 'Caprock thickness')
r = engine.evaluate_parameter(150, 'm', thickness)
assert r['pass'] and r['saw_score'] == 5, r

# 4. Exact requested hard-cutoff-only PASS behavior.
# 35.5 is >=20 but falls in none of 20–35, 36–50, ... ranges.
r = engine.evaluate_parameter(35.5, 'm', thickness)
assert r['pass'] is True and r['saw_score'] is None and r['weighted_score'] is None, r

# Qualitative hard-cutoff text is not one of the SAW texts for Trap type.
trap = next(x for x in reference['categories']['Technical'] if x['parameter'] == 'Trap type')
r = engine.evaluate_parameter('Structural/Stratigraphic', '', trap)
assert r['pass'] is True and r['saw_score'] is None and r['weighted_score'] is None, r

# Qualitative SAW text must receive its actual reference score.
r = engine.evaluate_parameter('Stratigraphic trap', '', trap)
assert r['pass'] is True and r['saw_score'] == 3 and r['weighted_score'] is not None, r

# 5. Construct a field using a valid SAW value for every reference parameter.
# This tests the entire screen_one_field aggregation and specifically verifies
# that no None weighted score is ever added to weighted_sum.
field = {'values': {c: {} for c in engine.CATEGORIES},
         'units': {c: {} for c in engine.CATEGORIES},
         'economics': {'actual capex': 1000000, 'actual opex': 50000}}

for cat in engine.CATEGORIES:
    for item in reference['categories'][cat]:
        ranges = item['saw_ranges']
        dtype = engine.norm_text(item['data_type'])
        if dtype == 'qualitative':
            value = next((a for a,b,_ in ranges if not engine.is_missing(a)), None)
        else:
            # Choose a representative value from the first SAW range.
            a, b, _ = ranges[0]
            av = engine.parse_number_strict(a)
            bv = engine.parse_number_strict(b)
            if av is not None and bv is not None:
                value = (av + bv) / 2
            elif av is not None:
                # Open-ended lower bound (>x / >=x), choose just above x.
                op, _ = engine.parse_operator_cutoff(a)
                value = av + max(abs(av) * 0.1, 1e-12) if op in ('>', '>=') else av / 2
            elif bv is not None:
                value = bv / 2
            else:
                raise AssertionError(f'No usable SAW reference for {item["parameter"]}')
        key = engine.norm_text(item['parameter'])
        field['values'][cat][key] = value
        field['units'][cat][key] = item['data_type']

result = engine.screen_one_field('Test Field', field, reference)
assert result['Passed'] is True, result['Reasons']
assert result['Overall SAW Score'] is not None
assert all(result[c] is not None for c in engine.CATEGORIES)

# 6. Verify missing input is treated as a field failure, not as a crash.
missing_field = {'values': {c: {} for c in engine.CATEGORIES},
                 'units': {c: {} for c in engine.CATEGORIES},
                 'economics': {'actual capex': 1000000, 'actual opex': 50000}}
r = engine.screen_one_field('Missing Field', missing_field, reference)
assert r['Passed'] is False and r['Reasons']

print('ALL CCS ENGINE REGRESSION TESTS PASSED')
