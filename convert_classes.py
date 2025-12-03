#!/usr/bin/env python3
import re

with open('ConsolidatedDashboard.gs', 'r') as f:
    lines = f.readlines()

output = []
i = 0
while i < len(lines):
    line = lines[i]

    # Convert DistributedLock class
    if 'class DistributedLock {' in line:
        # Skip the class declaration line
        i += 1
        # Copy comments until constructor
        while i < len(lines) and 'constructor(' not in lines[i]:
            output.append(lines[i])
            i += 1

        # Convert constructor
        if 'constructor(resource, timeout)' in lines[i]:
            output.append('function DistributedLock(resource, timeout) {\n')
            i += 1
            # Copy constructor body until closing brace
            brace_count = 1
            while i < len(lines) and brace_count > 0:
                if '{' in lines[i]:
                    brace_count += lines[i].count('{')
                if '}' in lines[i]:
                    brace_count -= lines[i].count('}')
                if brace_count > 0:
                    output.append(lines[i])
                elif brace_count == 0:
                    # Add default parameter handling
                    output.insert(-3, '  this.timeout = timeout !== undefined ? timeout : 30000;\n')
                    # Remove the original timeout line
                    for j in range(len(output)-1, len(output)-10, -1):
                        if 'this.timeout = timeout' in output[j]:
                            output.pop(j)
                            break
                    output.append('}\n')
                i += 1

            # Now convert methods to prototype methods
            while i < len(lines):
                # Check for method declaration
                method_match = re.match(r'^\s+(\w+)\s*\((.*?)\)\s*\{', lines[i])
                if method_match and not 'function' in lines[i]:
                    method_name = method_match.group(1)
                    params = method_match.group(2)

                    # Add blank line and JSDoc if present
                    output.append('\n')
                    j = i - 1
                    jsdoc = []
                    while j >= 0 and (lines[j].strip().startswith('*') or lines[j].strip().startswith('/**') or lines[j].strip() == ''):
                        if lines[j].strip():
                            jsdoc.insert(0, lines[j])
                        j -= 1
                    output.extend(jsdoc)

                    output.append(f'DistributedLock.prototype.{method_name} = function({params}) {{\n')
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
                        elif brace_count == 0:
                            output.append('};\n')
                        i += 1
                    continue

                # Check for end of class
                if re.match(r'^\}', lines[i]) and i > 11800:
                    # Don't output the closing brace
                    i += 1
                    break

                output.append(lines[i])
                i += 1
        continue

    # Convert Transaction class similarly
    if 'class Transaction {' in line:
        i += 1
        while i < len(lines) and 'constructor(' not in lines[i]:
            output.append(lines[i])
            i += 1

        if 'constructor(spreadsheet)' in lines[i]:
            output.append('function Transaction(spreadsheet) {\n')
            i += 1
            brace_count = 1
            while i < len(lines) and brace_count > 0:
                if '{' in lines[i]:
                    brace_count += lines[i].count('{')
                if '}' in lines[i]:
                    brace_count -= lines[i].count('}')
                if brace_count > 0:
                    output.append(lines[i])
                elif brace_count == 0:
                    output.append('}\n')
                i += 1

            # Convert methods
            while i < len(lines):
                method_match = re.match(r'^\s+(\w+)\s*\((.*?)\)\s*\{', lines[i])
                if method_match and not 'function' in lines[i]:
                    method_name = method_match.group(1)
                    params = method_match.group(2)

                    output.append('\n')
                    j = i - 1
                    jsdoc = []
                    while j >= 0 and (lines[j].strip().startswith('*') or lines[j].strip().startswith('/**') or lines[j].strip() == ''):
                        if lines[j].strip():
                            jsdoc.insert(0, lines[j])
                        j -= 1
                    output.extend(jsdoc)

                    output.append(f'Transaction.prototype.{method_name} = function({params}) {{\n')
                    i += 1

                    brace_count = 1
                    while i < len(lines) and brace_count > 0:
                        if '{' in lines[i]:
                            brace_count += lines[i].count('{')
                        if '}' in lines[i]:
                            brace_count -= lines[i].count('}')
                        if brace_count > 0:
                            output.append(lines[i])
                        elif brace_count == 0:
                            output.append('};\n')
                        i += 1
                    continue

                if re.match(r'^\}', lines[i]) and i > 27000:
                    i += 1
                    break

                output.append(lines[i])
                i += 1
        continue

    output.append(line)
    i += 1

with open('ConsolidatedDashboard.gs', 'w') as f:
    f.writelines(output)

print("Converted classes to ES5")
