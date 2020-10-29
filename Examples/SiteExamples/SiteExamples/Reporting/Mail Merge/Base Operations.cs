using System.Data;
using System.Data.OleDb;
using Aspose.Words;
using NUnit.Framework;

namespace SiteExamples.Reporting.Mail_Merge
{
    class BaseOperations : SiteExamplesBase
    {
        [Test]
        public static void SimpleMailMerge()
        {
            //ExStart:SimpleMailMerge
            Document doc = new Document(MyDir + "Mail merge destinations - Complex template.docx");

            doc.MailMerge.UseNonMergeFields = true;

            // Fill the fields in the document with user data
            doc.MailMerge.Execute(
                new string[] { "FullName", "Company", "Address", "Address2", "City" },
                new object[] { "James Bond", "MI5 Headquarters", "Milbank", "", "London" });

            // Send the document in Word format to the client browser with an option to save to disk or open inside the current browser
            doc.Save(ArtifactsDir + "SimpleMailMerge.docx");
            //ExEnd:SimpleMailMerge
        }

        [Test]
        public static void UseIfElseMustacheSyntax()
        {
            //ExStart:UseOfifelseMustacheSyntax
            Document doc = new Document(MyDir + "Mail merge destinations - Mustache syntax.docx");

            doc.MailMerge.UseNonMergeFields = true;
            doc.MailMerge.Execute(new string[] { "GENDER" }, new object[] { "MALE" });

            doc.Save(ArtifactsDir + "MailMerge.IfElseMustacheSyntax.docx");
            //ExEnd:UseOfifelseMustacheSyntax
        }

        [Test]
        public static void ExecuteWithRegionsDataTable()
        {
            //ExStart:ExecuteWithRegionsDataTable
            Document doc = new Document(MyDir + "Mail merge destinations - Orders.docx");

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
                                MyDir + "Northwind.mdb";
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

        [Test]
        public static void ProduceMultipleDocuments()
        {
            //ExStart:ProduceMultipleDocuments
            string connString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + MyDir + "Mail merge data - Customers.mdb";
            
            OleDbConnection conn = new OleDbConnection(connString);
            conn.Open();
            // Get data from a database
            OleDbCommand cmd = new OleDbCommand("SELECT * FROM Customers", conn);
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            
            DataTable data = new DataTable();
            da.Fill(data);

            // Open the template document
            Document doc = new Document(MyDir + "Mail merge destinations - Northwind traders.docx");

            int counter = 1;
            // Loop though all records in the data source
            foreach (DataRow row in data.Rows)
            {
                // Clone the template instead of loading it from disk (for speed)
                Document dstDoc = (Document) doc.Clone(true);

                // Execute mail merge
                dstDoc.MailMerge.Execute(row);

                // Save the document
                dstDoc.Save(string.Format(ArtifactsDir + "MailMerge.ProduceMultipleDocuments_{0}.doc", counter++));
            }
            //ExEnd:ProduceMultipleDocuments
        }
    }
}