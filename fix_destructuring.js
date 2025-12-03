// Node.js script to fix ES6 destructuring in function parameters
const fs = require('fs');

const filePath = './ConsolidatedDashboard.gs';
let content = fs.readFileSync(filePath, 'utf8');

// Fix pattern: .filter(function([param]) { ... })
// Replace with: .filter(function(item) { var param = item[0]; ... })
content = content.replace(
  /\.filter\(function\(\[(\w+)\]\)\s*\{/g,
  '.filter(function(item) { var $1 = item[0];'
);

// Fix pattern: .map(function([a, b]) { return [a, b]; })
// This is a simple identity map, can be simplified
content = content.replace(
  /\.map\(function\(\[(\w+),\s*(\w+)\]\)\s*\{\s*return\s*\[\1,\s*\2\];\s*\}\)/g,
  '.map(function(item) { return [item[0], item[1]]; })'
);

// Fix pattern: .map(function([a, b]) { ... })
// Replace with: .map(function(item) { var a = item[0]; var b = item[1]; ... })
content = content.replace(
  /\.map\(function\(\[(\w+),\s*(\w+)\]\)\s*\{/g,
  '.map(function(item) { var $1 = item[0]; var $2 = item[1];'
);

// Fix pattern: .forEach(function([a, b]) { ... })
content = content.replace(
  /\.forEach\(function\(\[(\w+),\s*(\w+)\]\)\s*\{/g,
  '.forEach(function(entry) { var $1 = entry[0]; var $2 = entry[1];'
);

fs.writeFileSync(filePath, content, 'utf8');
console.log('Fixed destructuring patterns');
