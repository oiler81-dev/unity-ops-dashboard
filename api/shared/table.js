const { TableClient } = require("@azure/data-tables");

function getConnectionString() {
  const value = process.env.AZURE_TABLES_CONNECTION_STRING;
  if (!value) {
    throw new Error("Missing AZURE_TABLES_CONNECTION_STRING");
  }
  return value;
}

function getTableClient(tableName) {
  return TableClient.fromConnectionString(getConnectionString(), tableName);
}

async function ensureTable(tableName) {
  const client = getTableClient(tableName);

  try {
    await client.createTable();
  } catch (error) {
    if (error?.statusCode !== 409) {
      throw error;
    }
  }

  return client;
}

module.exports = {
  getTableClient,
  ensureTable
};
