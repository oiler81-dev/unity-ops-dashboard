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
    throw new Error(
      "Missing AZURE_STORAGE_CONNECTION_STRING or AzureWebJobsStorage"
    );
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

    async upsertEntity(entity) {
      await this.ensureTable();
      return client.upsertEntity(entity, "Merge");
    },

    async *listEntities() {
      await this.ensureTable();
      for await (const entity of client.listEntities()) {
        yield entity;
      }
    }
  };
}

module.exports = { getTableClient };
