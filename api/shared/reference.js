const { getTableClient } = require("./table");

async function getReferenceMapForEntity(entity) {
  const targetsTable = getTableClient("ReferenceTargets");
  const thresholdsTable = getTableClient("ReferenceThresholds");

  const [targetRows, thresholdRows] = await Promise.all([
    targetsTable.listByPartition(entity),
    thresholdsTable.listByPartition(entity)
  ]);

  const map = {};

  for (const row of targetRows) {
    const key = row.rowKey || row.RowKey;
    if (!key) continue;
    if (!map[key]) map[key] = {};
    map[key].label = row.label || key;
    map[key].targetValue = row.targetValue ?? null;
  }

  for (const row of thresholdRows) {
    const key = row.rowKey || row.RowKey;
    if (!key) continue;
    if (!map[key]) map[key] = {};
    map[key].threshold = {
      comparisonType: row.comparisonType || "",
      greenMin: row.greenMin ?? null,
      yellowMin: row.yellowMin ?? null,
      redMin: row.redMin ?? null,
      greenMax: row.greenMax ?? null,
      yellowMax: row.yellowMax ?? null,
      redMax: row.redMax ?? null
    };
  }

  return map;
}

module.exports = { getReferenceMapForEntity };
