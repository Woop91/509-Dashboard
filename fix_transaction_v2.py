#!/usr/bin/env python3
"""
Properly convert Transaction class to ES5
"""
import re

with open('ConsolidatedDashboard.gs', 'r') as f:
    content = f.read()

# Find Transaction function and extract the entire class
pattern = r'(function Transaction\(spreadsheet\) \{\s*this\.ss = spreadsheet;\s*this\.snapshots = \{\};[^}]*?this\.rolledBack = false;\s*)\n\s*/\*\*'

# Replace the pattern to close the constructor before methods
replacement = r'\1}\n\n/**'

content = re.sub(pattern, replacement, content, flags=re.DOTALL)

# Now convert ES6 methods to prototype methods in the Transaction section
# Find all methods after Transaction constructor

# Pattern: '  snapshot(sheetName) {' -> 'Transaction.prototype.snapshot = function(sheetName) {'
content = re.sub(
    r'(\n/\*\*[^*]*?\*/)\s+(\w+)\s*\(([^)]*)\)\s*\{',
    lambda m: f'{m.group(1)}\nTransaction.prototype.{m.group(2)} = function({m.group(3)}) {{',
    content[content.find('function Transaction'):content.find('function Transaction') + 5000]
)

# Find and remove the closing brace of Transaction class
# Look for standalone '}' after Transaction methods
transaction_end_pattern = r'(Transaction\.prototype\.\w+[^}]*\};\s*)\n\}'
content = re.sub(transaction_end_pattern, r'\1', content)

with open('ConsolidatedDashboard.gs', 'w') as f:
    f.write(content)

print("Transaction class fixed!")
