import ast
import pandas as pd
import textwrap
from pathlib import Path

# Load monthly_recap.py source
src_path = Path(__file__).resolve().parents[1] / 'monthly_recap.py'
src = src_path.read_text()

# Parse AST and extract the extract_dealership_name function source
mod = ast.parse(src)
func_node = None
for node in mod.body:
    if isinstance(node, ast.FunctionDef) and node.name == 'extract_dealership_name':
        func_node = node
        break

if func_node is None:
    raise RuntimeError('extract_dealership_name not found in monthly_recap.py')

func_src = ast.get_source_segment(src, func_node)
# Ensure any relative imports or dependencies are available; the function uses only pandas and re
# Build an isolated namespace and exec the function definition
ns = {'pd': pd, 're': __import__('re')}
exec(func_src, ns)
extract_dealership_name = ns['extract_dealership_name']

# Build synthetic DataFrame matching the expected Excel structure
data = [
    {'Dealership Name':'Hendrick BMW','First Name':'Chris','Last Name':'Maroulis','Learning Days':3,'Total Learning Units':127},
    {'Dealership Name':'Hendrick BMW','First Name':'Xavier','Last Name':'Wallace','Learning Days':1,'Total Learning Units':69},
    {'Dealership Name':'Hendrick Audi','First Name':'Alex','Last Name':'Smith','Learning Days':5,'Total Learning Units':40},
    {'Dealership Name':'Hendrick Audi','First Name':'Jamie','Last Name':'Carpenter','Learning Days':2,'Total Learning Units':22},
]

df = pd.DataFrame(data)

# Call the extracted function
mapping = extract_dealership_name(df, 'test.xlsx')

# Basic assertions
assert 'Hendrick BMW' in mapping, 'Hendrick BMW not found in mapping'
assert 'Hendrick Audi' in mapping, 'Hendrick Audi not found in mapping'
assert len(mapping['Hendrick BMW']) == 2, f"Expected 2 employees for Hendrick BMW, got {len(mapping['Hendrick BMW'])}"
assert mapping['Hendrick BMW'][0]['First Name'] == 'Chris'

print('Mapping result:')
for dealer, emps in mapping.items():
    print(f"Dealer: {dealer}")
    for e in emps:
        print(f"  - {e['First Name']} {e['Last Name']}: days={e['Learning Days']}, units={e['Total Learning Units']}")

print('\nAll tests passed.')
