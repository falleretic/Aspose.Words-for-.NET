﻿using System.IO;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class ConvertDocumentToByte : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:ConvertDocumentToByte
            Document doc = new Document(LoadingSavingDir + "Test File (doc).doc");

            // Create a new memory stream
            MemoryStream outStream = new MemoryStream();
            // Save the document to stream
            doc.Save(outStream, SaveFormat.Docx);

            // Convert the document to byte form
            byte[] docBytes = outStream.ToArray();

            // The bytes are now ready to be stored/transmitted
            // Now reverse the steps to load the bytes back into a document object
            MemoryStream inStream = new MemoryStream(docBytes);

            // Load the stream into a new document object
            Document loadDoc = new Document(inStream);
            //ExEnd:ConvertDocumentToByte
        }
    }
}