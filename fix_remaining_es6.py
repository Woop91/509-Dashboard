#!/usr/bin/env python3
"""
Comprehensive ES6 to ES5 conversion for Rhino compatibility
"""
import re

with open('ConsolidatedDashboard.gs', 'r') as f:
    content = f.read()

print("Starting comprehensive ES6 to ES5 conversion...")

# Fix 1: Convert arrow functions to regular functions
# Pattern: .map(([a, b]) => [...])
content = re.sub(
    r'\.map\(\(\[(\w+), (\w+)\]\) => \[',
    r'.map(function(item) { var \1 = item[0]; var \2 = item[1]; return [',
    content
)

# Pattern: .map((item, index) => `template`)
content = re.sub(
    r'\.map\(\((\w+)(?:, (\w+))?\) => `',
    lambda m: f'.map(function({m.group(1)}{", " + m.group(2) if m.group(2) else ""}) {{ return `',
    content
)

# Pattern: .map(member => `template`)
content = re.sub(
    r'\.map\((\w+) => `',
    r'.map(function(\1) { return `',
    content
)

# Fix 2: Fix destructuring in function parameters
# Pattern: function([type, total], index)
content = re.sub(
    r'function\(\[(\w+), (\w+)\], (\w+)\)',
    r'function(item, \3) { var \1 = item[0]; var \2 = item[1];',
    content
)

# Fix 3: Replace new Set() with object-based implementation
# Pattern: new Set()
set_counter = 0
def replace_set(match):
    global set_counter
    set_counter += 1
    return '{}'  # Use object as set

content = re.sub(r'new Set\(\)', replace_set, content)

# Fix 4: Fix Set operations
# setObject.add(value) -> setObject[value] = true
content = re.sub(r'(\w+)\.add\(([^)]+)\)', r'\1[\2] = true', content)
# setObject.has(value) -> (value in setObject)
content = re.sub(r'(\w+)\.has\(([^)]+)\)', r'(\2 in \1)', content)
# setObject.size -> Object.keys(setObject).length
content = re.sub(r'(\w+)\.size([^a-zA-Z_])', r'Object.keys(\1).length\2', content)

print(f"Fixed {set_counter} Set() instances")

with open('ConsolidatedDashboard.gs', 'w') as f:
    f.write(content)

print("Conversion complete!")
print("- Fixed arrow functions")
print("- Fixed destructuring parameters")
print("- Replaced new Set() with objects")
