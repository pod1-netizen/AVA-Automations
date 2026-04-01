f = open('src/App.jsx', 'r')
c = f.read()
f.close()

old = """function buildClients(sheetData) {
  // Build a normalized lookup: canonical display name \u2192 CSV key
  const csvKeyMap = {};
  Object.keys(sheetData).forEach(csvKey => {
    const canonical = CLIENT_ALIASES[csvKey] !== undefined ? CLIENT_ALIASES[csvKey] : csvKey;
    if (canonical) csvKeyMap[canonical] = csvKey;
  });
  return CLIENT_META.map(m => ({
    ...m,
    dataKey: csvKeyMap[m.display] || null,
  }));
}"""

new = """function buildClients(sheetData) {
  // Build a normalized lookup: canonical display name \u2192 array of CSV keys
  const csvKeyMap = {};
  Object.keys(sheetData).forEach(csvKey => {
    const canonical = CLIENT_ALIASES[csvKey] !== undefined ? CLIENT_ALIASES[csvKey] : csvKey;
    if (canonical) {
      if (!csvKeyMap[canonical]) csvKeyMap[canonical] = [];
      csvKeyMap[canonical].push(csvKey);
    }
  });
  return CLIENT_META.map(m => ({
    ...m,
    dataKey: csvKeyMap[m.display] ? csvKeyMap[m.display][0] : null,
    dataKeys: csvKeyMap[m.display] || [],
  }));
}"""

c = c.replace(old, new)
f = open('src/App.jsx', 'w')
f.write(c)
f.close()
print('Done!')
