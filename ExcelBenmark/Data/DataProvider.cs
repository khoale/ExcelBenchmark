namespace ExcelBenmark.Data
{
    using System.Collections.Generic;
    using System.Configuration;
    using System.Data.SqlClient;

    using Dapper;

    public class DataProvider
    {
        private SqlConnection connection;

        public DataProvider()
        {
            var connectionString = ConfigurationManager.ConnectionStrings["DbConnection"].ConnectionString;
            this.connection = new SqlConnection(connectionString);
        }

        public IEnumerable<dynamic> GetProducts()
        {
            return this.connection.Query("SELECT *FROM [Production].[Product]");
        } 
    }
}