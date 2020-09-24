using System.Data;
using System.Data.OleDb;
using System.IO;
using Aspose.Words.MailMerging;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class MailMergeImageFromBlob : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:MailMergeImageFromBlob
            Document doc = new Document(MailMergeDir + "Mail merge destination - Northwind employees.docx");

            // Set up the event handler for image fields
            doc.MailMerge.FieldMergingCallback = new HandleMergeImageFieldFromBlob();

            // Open a database connection
            string connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + DatabaseDir + "Northwind.mdb";
            OleDbConnection conn = new OleDbConnection(connString);
            conn.Open();

            // Open the data reader
            // It needs to be in the normal mode that reads all record at once
            OleDbCommand cmd = new OleDbCommand("SELECT * FROM Employees", conn);
            IDataReader dataReader = cmd.ExecuteReader();

            // Perform mail merge
            doc.MailMerge.ExecuteWithRegions(dataReader, "Employees");

            // Close the database
            conn.Close();
            
            doc.Save(ArtifactsDir + "MailMerge.ImageFromBlob.docx");
            //ExEnd:MailMergeImageFromBlob
        }

        //ExStart:HandleMergeImageFieldFromBlob 
        public class HandleMergeImageFieldFromBlob : IFieldMergingCallback
        {
            void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
            {
                // Do nothing
            }

            /// <summary>
            /// This is called when mail merge engine encounters Image:XXX merge field in the document.
            /// You have a chance to return an Image object, file name or a stream that contains the image.
            /// </summary>
            void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs e)
            {
                // The field value is a byte array, just cast it and create a stream on it
                MemoryStream imageStream = new MemoryStream((byte[]) e.FieldValue);
                // Now the mail merge engine will retrieve the image from the stream
                e.ImageStream = imageStream;
            }
        }
        //ExEnd:HandleMergeImageFieldFromBlob
    }
}