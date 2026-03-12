// api/shared/table.js
const { TableClient } = require("@azure/data-tables");

let tableClientFactory = null;

function getConnectionString() {
  return process.env.AZURE_STORAGE_CONNECTION_STRING || process.env.AzureWebJobsStorage || null;
}

function createTableClient(tableName) {
  const conn = getConnectionString();
  if (!conn) return null;
  return TableClient.fromConnectionString(conn, tableName);
}

module.exports = {
  getTableClient: (tableName) => {
    try {
      return createTableClient(tableName);
    } catch (err) {
      // don't crash at import-time; surface null to caller
      return null;
    }
  },
  ensureTable: async (tableName) => {
    const client = createTableClient(tableName);
    if (!client) throw new Error("Azure storage connection string not configured");
    await client.createTable();
    return client;
  },
  upsertEntity: async (tableName, entity) => {
    const client = createTableClient(tableName);
    if (!client) throw new Error("Azure storage connection string not configured");
    return client.upsertEntity(entity);
  },
  getEntity: async (tableName, partitionKey, rowKey) => {
    const client = createTableClient(tableName);
    if (!client) throw new Error("Azure storage connection string not configured");
    return client.getEntity(partitionKey, rowKey);
  },
  listEntities: async (tableName, options) => {
    const client = createTableClient(tableName);
    if (!client) throw new Error("Azure storage connection string not configured");
    const entities = [];
    for await (const e of client.listEntities(options || {})) {
      entities.push(e);
    }
    return entities;
  }
};
