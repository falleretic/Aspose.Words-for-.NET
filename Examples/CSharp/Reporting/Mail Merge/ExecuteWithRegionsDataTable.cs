using System.Data;
using System.Data.OleDb;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class ExecuteWithRegionsDataTable : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:ExecuteWithRegionsDataTable
            Document doc = new Document(MailMergeDir + "Mail merge destinations - Orders.docx");

            const int orderId = 10444;

            // Perform several mail merge operations populating only part of the document each time

            // Use DataTable as a data source
            DataTable orderTable = GetTestOrder(orderId);
            doc.MailMerge.ExecuteWithRegions(orderTable);

            // Instead of using DataTable you can create a DataView for custom sort or filter and then mail merge
            DataView orderDetailsView = new DataView(GetTestOrderDetails(orderId));
            orderDetailsView.Sort = "ExtendedPrice DESC";
 
            // Execute the mail merge operation.
            doc.MailMerge.ExecuteWithRegions(orderDetailsView);

            doc.Save(ArtifactsDir + "MailMerge.ExecuteWithRegions.docx");
            //ExEnd:ExecuteWithRegionsDataTable
        }

        //ExStart:ExecuteWithRegionsDataTableMethods
        private static DataTable GetTestOrder(int orderId)
        {
            DataTable table = ExecuteDataTable($"SELECT * FROM AsposeWordOrders WHERE OrderId = {orderId}");
            table.TableName = "Orders";
            
            return table;
        }

        private static DataTable GetTestOrderDetails(int orderId)
        {
            DataTable table = ExecuteDataTable(
                $"SELECT * FROM AsposeWordOrderDetails WHERE OrderId = {orderId} ORDER BY ProductID");
            table.TableName = "OrderDetails";
            
            return table;
        }

        /// <summary>
        /// Utility function that creates a connection, command, 
        /// Executes the command and return the result in a DataTable.
        /// </summary>
        private static DataTable ExecuteDataTable(string commandText)
        {
            // Open the database connection.
            string connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +
                                DatabaseDir + "Northwind.mdb";
            OleDbConnection conn = new OleDbConnection(connString);
            conn.Open();

            // Create and execute a command
            OleDbCommand cmd = new OleDbCommand(commandText, conn);
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataTable table = new DataTable();
            da.Fill(table);

            // Close the database
            conn.Close();

            return table;
        }
        //ExEnd:ExecuteWithRegionsDataTableMethods
    }
}