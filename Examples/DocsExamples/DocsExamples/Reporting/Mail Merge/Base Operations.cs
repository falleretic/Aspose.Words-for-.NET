using System.Data;
using System.Data.OleDb;
using Aspose.Words;
using NUnit.Framework;

namespace DocsExamples.Reporting.Mail_Merge
{
    internal class BaseOperations : DocsExamplesBase
    {
        [Test]
        public static void SimpleMailMerge()
        {
            //ExStart:SimpleMailMerge
            Document doc = new Document(MyDir + "Mail merge destinations - Complex template.docx");

            doc.MailMerge.UseNonMergeFields = true;

            doc.MailMerge.Execute(
                new[] { "FullName", "Company", "Address", "Address2", "City" },
                new object[] { "James Bond", "MI5 Headquarters", "Milbank", "", "London" });

            doc.Save(ArtifactsDir + "BaseOperations.SimpleMailMerge.docx");
            //ExEnd:SimpleMailMerge
        }

        [Test]
        public static void UseIfElseMustache()
        {
            //ExStart:UseOfifelseMustacheSyntax
            Document doc = new Document(MyDir + "Mail merge destinations - Mustache syntax.docx");

            doc.MailMerge.UseNonMergeFields = true;
            doc.MailMerge.Execute(new[] { "GENDER" }, new object[] { "MALE" });

            doc.Save(ArtifactsDir + "BaseOperations.IfElseMustache.docx");
            //ExEnd:UseOfifelseMustacheSyntax
        }

        [Test]
        public static void ExecuteWithRegionsDataTable()
        {
            //ExStart:ExecuteWithRegionsDataTable
            Document doc = new Document(MyDir + "Mail merge destinations - Orders.docx");

            const int orderId = 10444;

            // Perform several mail merge operations populating only part of the document each time.

            DataTable orderTable = GetTestOrder(orderId);
            doc.MailMerge.ExecuteWithRegions(orderTable);

            // Instead of using DataTable you can create a DataView for custom sort or filter and then mail merge.
            DataView orderDetailsView = new DataView(GetTestOrderDetails(orderId));
            orderDetailsView.Sort = "ExtendedPrice DESC";
 
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
        /// Utility function that creates a connection, command, executes the command and returns the result in a DataTable.
        /// </summary>
        private static DataTable ExecuteDataTable(string commandText)
        {
            string connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + DatabaseDir + "Northwind.mdb";

            OleDbConnection conn = new OleDbConnection(connString);
            conn.Open();

            OleDbCommand cmd = new OleDbCommand(commandText, conn);
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);

            DataTable table = new DataTable();
            da.Fill(table);

            conn.Close();

            return table;
        }
        //ExEnd:ExecuteWithRegionsDataTableMethods

        [Test]
        public static void ProduceMultipleDocuments()
        {
            //ExStart:ProduceMultipleDocuments
            string connString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + DatabaseDir + "Northwind.mdb";
            
            OleDbConnection conn = new OleDbConnection(connString);
            conn.Open();
            
            OleDbCommand cmd = new OleDbCommand("SELECT * FROM Customers", conn);
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            
            DataTable data = new DataTable();
            da.Fill(data);

            Document doc = new Document(MyDir + "Mail merge destination - Northwind suppliers.docx");

            int counter = 1;
            foreach (DataRow row in data.Rows)
            {
                Document dstDoc = (Document) doc.Clone(true);

                dstDoc.MailMerge.Execute(row);

                dstDoc.Save(string.Format(ArtifactsDir + "BaseOperations.ProduceMultipleDocuments_{0}.docx", counter++));
            }
            //ExEnd:ProduceMultipleDocuments
        }
    }
}