using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class LoadAndSaveDocToDatabaseEx : TestDataHelper
    {
        [Test]
        public static void LoadAndSaveDocToDatabase()
        {
            Document doc = new Document(LoadingSavingDir + "Document.docx");
            
            //ExStart:OpenDatabaseConnection 
            // Open a database connection
            string connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + DatabaseDir + "Northwind.mdb";
            OleDbConnection connection = new OleDbConnection(connString);
            connection.Open();
            //ExEnd:OpenDatabaseConnection
            
            //ExStart:OpenRetrieveAndDelete 
            // Store the document to the database
            StoreToDatabase(doc, connection);
            
            // Read the document from the database and store the file to disk
            Document dbDoc = ReadFromDatabase("Document.docx", connection);
            
            // Save the retrieved document to disk
            dbDoc.Save(ArtifactsDir + "LoadAndSaveDocToDatabase.docx");

            // Delete the document from the database
            DeleteFromDatabase("Document.docx", connection);

            // Close the connection to the database
            connection.Close();
            //ExEnd:OpenRetrieveAndDelete 
        }

        //ExStart:StoreToDatabase 
        public static void StoreToDatabase(Document doc, OleDbConnection connection)
        {
            // Save the document to a MemoryStream object
            MemoryStream stream = new MemoryStream();
            doc.Save(stream, SaveFormat.Docx);

            // Get the filename from the document
            string fileName = Path.GetFileName(doc.OriginalFileName);

            // Create the SQL command
            string commandString = "INSERT INTO Documents (Name, Data) VALUES('" + fileName + "', @Doc)";
            OleDbCommand command = new OleDbCommand(commandString, connection);

            // Add the @Doc parameter
            command.Parameters.AddWithValue("Doc", stream.ToArray());

            // Write the document to the database
            command.ExecuteNonQuery();
        }
        //ExEnd:StoreToDatabase
        
        //ExStart:ReadFromDatabase 
        public static Document ReadFromDatabase(string fileName, OleDbConnection connection)
        {
            // Create the SQL command
            string commandString = "SELECT * FROM Documents WHERE Name='" + fileName + "'";
            OleDbCommand command = new OleDbCommand(commandString, connection);

            // Create the data adapter
            OleDbDataAdapter adapter = new OleDbDataAdapter(command);

            // Fill the results from the database into a DataTable
            DataTable dataTable = new DataTable();
            adapter.Fill(dataTable);

            // Check there was a matching record found from the database and throw an exception if no record was found
            if (dataTable.Rows.Count == 0)
                throw new ArgumentException(
                    $"Could not find any record matching the document \"{fileName}\" in the database.");

            // The document is stored in byte form in the FileContent column
            // Retrieve these bytes of the first matching record to a new buffer
            byte[] buffer = (byte[]) dataTable.Rows[0]["Data"];

            // Wrap the bytes from the buffer into a new MemoryStream object
            MemoryStream newStream = new MemoryStream(buffer);

            // Read the document from the stream
            Document doc = new Document(newStream);

            // Return the retrieved document
            return doc;
        }
        //ExEnd:ReadFromDatabase
        
        //ExStart:DeleteFromDatabase 
        public static void DeleteFromDatabase(string fileName, OleDbConnection connection)
        {
            // Create the SQL command
            string commandString = "DELETE * FROM Documents WHERE Name='" + fileName + "'";
            OleDbCommand command = new OleDbCommand(commandString, connection);

            // Delete the record
            command.ExecuteNonQuery();
        }
        //ExEnd:DeleteFromDatabase
    }
}