const { TableClient } = require("@azure/data-tables");

function getConnectionString() {
  return (
    process.env.AZURE_STORAGE_CONNECTION_STRING ||
    process.env.AzureWebJobsStorage ||
    ""
  );
}

function getTableClient(tableName) {
  const connectionString = getConnectionString();

  if (!connectionString) {
    throw new Error("Missing AZURE_STORAGE_CONNECTION_STRING or AzureWebJobsStorage.");
  }

  const client = TableClient.fromConnectionString(connectionString, tableName);

  return {
    async ensureTable() {
      try {
        await client.createTable();
      } catch (err) {
        if (err.statusCode !== 409) {
          throw err;
        }
      }
    },

    async upsertEntity(entity, mode = "Merge") {
      await this.ensureTable();
      return client.upsertEntity(entity, mode);
    },

    async getEntity(partitionKey, rowKey) {
      await this.ensureTable();
      return client.getEntity(partitionKey, rowKey);
    },

    async deleteEntity(partitionKey, rowKey) {
      await this.ensureTable();
      return client.deleteEntity(partitionKey, rowKey);
    },

    async listByPartitionKey(partitionKey) {
      await this.ensureTable();
      const results = [];
      const filter = `PartitionKey eq '${String(partitionKey).replace(/'/g, "''")}'`;

      for await (const entity of client.listEntities({ queryOptions: { filter } })) {
        results.push(entity);
      }

      return results;
    },

    async query(filter) {
      await this.ensureTable();
      const results = [];

      for await (const entity of client.listEntities({ queryOptions: { filter } })) {
        results.push(entity);
      }

      return results;
    },

    async *listEntities(options = {}) {
      await this.ensureTable();
      for await (const entity of client.listEntities(options)) {
        yield entity;
      }
    }
  };
}

module.exports = { getTableClient };
