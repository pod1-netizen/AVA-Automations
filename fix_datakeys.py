f = open('src/App.jsx', 'r')
c = f.read()
f.close()

old = """    if (!client?.dataKey || !period) return null;
    const raw = sheetData[client.dataKey];
    return raw ? filterData(raw, period) : null;"""

new = """    if (!client || !period) return null;
    const keys = client.dataKeys && client.dataKeys.length > 0 ? client.dataKeys : (client.dataKey ? [client.dataKey] : []);
    if (keys.length === 0) return null;
    const merged = { tasks: [], wins: [] };
    keys.forEach(k => {
      if (sheetData[k]) {
        merged.tasks = merged.tasks.concat(sheetData[k].tasks || []);
        merged.wins = merged.wins.concat(sheetData[k].wins || []);
      }
    });
    return merged.tasks.length > 0 ? filterData(merged, period) : null;"""

c = c.replace(old, new)
f = open('src/App.jsx', 'w')
f.write(c)
f.close()
print('Done!')
