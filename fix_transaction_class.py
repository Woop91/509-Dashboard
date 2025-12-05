#!/usr/bin/env python3
"""
Convert Transaction class methods to ES5 prototype methods
"""
import re

with open('ConsolidatedDashboard.gs', 'r') as f:
    lines = f.readlines()

output = []
i = 0
in_transaction = False
transaction_start = 0

while i < len(lines):
    line = lines[i]

    # Find Transaction constructor
    if 'function Transaction(spreadsheet)' in line:
        in_transaction = True
        transaction_start = i
        output.append(line)
        i += 1

        # Copy constructor body until we find a method
        brace_count = 1
        while i < len(lines):
            if '{' in lines[i]:
                brace_count += lines[i].count('{')
            if '}' in lines[i]:
                brace_count -= lines[i].count('}')

            # Check if this is a method declaration (not closing brace)
            if re.match(r'^\s+\w+\s*\([^)]*\)\s*\{', lines[i]) and brace_count == 1:
                # Found a method, need to close constructor and start converting methods
                output.append('}\n')
                break

            output.append(lines[i])
            i += 1

        # Now convert all methods
        while i < len(lines):
            # Look for method declaration
            method_match = re.match(r'^\s+(\w+)\s*\(([^)]*)\)\s*\{', lines[i])
            if method_match:
                method_name = method_match.group(1)
                params = method_match.group(2)

                # Collect JSDoc comments before the method
                jsdoc_lines = []
                j = i - 1
                while j >= transaction_start and (lines[j].strip().startswith('*') or lines[j].strip().startswith('/**') or lines[j].strip() == ''):
                    if lines[j].strip():
                        jsdoc_lines.insert(0, lines[j])
                    j -= 1

                # Write JSDoc
                output.append('\n')
                output.extend(jsdoc_lines)

                # Write prototype method
                output.append(f'Transaction.prototype.{method_name} = function({params}) {{\n')
                i += 1

                # Copy method body
                brace_count = 1
                while i < len(lines) and brace_count > 0:
                    if '{' in lines[i]:
                        brace_count += lines[i].count('{')
                    if '}' in lines[i]:
                        brace_count -= lines[i].count('}')

                    if brace_count > 0:
                        output.append(lines[i])
                    else:
                        output.append('};\n')
                    i += 1
                continue

            # Check for end of Transaction class
            if re.match(r'^\}', lines[i]) and i > transaction_start + 20:
                # Skip the class closing brace
                i += 1
                in_transaction = False
                break

            # Skip JSDoc that will be added with methods
            if not (lines[i].strip().startswith('*') or lines[i].strip().startswith('/**')):
                output.append(lines[i])
            i += 1
        continue

    output.append(line)
    i += 1

with open('ConsolidatedDashboard.gs', 'w') as f:
    f.writelines(output)

print("Transaction class converted to ES5 prototype methods!")
