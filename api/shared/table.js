const { TableClient } = require("@azure/data-tables");

function getConnectionString() {
  const value =
    process.env.AZURE_STORAGE_CONNECTION_STRING ||
    process.env.AzureWebJobsStorage;

  if (!value) {
    throw new Error("Missing AZURE_STORAGE_CONNECTION_STRING or AzureWebJobsStorage.");
  }

  return value;
}

function getTableClient(tableName) {
  const client = TableClient.fromConnectionString(getConnectionString(), tableName);

  return {
    async ensureTable() {
      await client.createTable().catch(() => {});
    },

    async getEntity(partitionKey, rowKey) {
      await client.createTable().catch(() => {});
      try {
        return await client.getEntity(partitionKey, rowKey);
      } catch (err) {
        if (err.statusCode === 404) return null;
        throw err;
      }
    },

    async upsertEntity(entity, mode = "Merge") {
      await client.createTable().catch(() => {});
      return client.upsertEntity(entity, mode);
    },

    async createEntity(entity) {
      await client.createTable().catch(() => {});
      return client.createEntity(entity);
    },

    async listByPartition(partitionKey) {
      await client.createTable().catch(() => {});
      const items = [];
      for await (const entity of client.listEntities({
        queryOptions: { filter: `PartitionKey eq '${partitionKey}'` }
      })) {
        items.push(entity);
      }
      return items;
    },

    async listAll() {
      await client.createTable().catch(() => {});
      const items = [];
      for await (const entity of client.listEntities()) {
        items.push(entity);
      }
      return items;
    }
  };
}

module.exports = { getTableClient };
