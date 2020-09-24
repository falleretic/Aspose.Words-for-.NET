using System.Data;
using System.Data.OleDb;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class ProduceMultipleDocuments : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:ProduceMultipleDocuments
            string connString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + MailMergeDir + "Mail merge data - Customers.mdb";
            
            OleDbConnection conn = new OleDbConnection(connString);
            conn.Open();
            // Get data from a database
            OleDbCommand cmd = new OleDbCommand("SELECT * FROM Customers", conn);
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            
            DataTable data = new DataTable();
            da.Fill(data);

            // Open the template document
            Document doc = new Document(MailMergeDir + "Mail merge destinations - Northwind traders.docx");

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