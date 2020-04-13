using System.IO;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class LoadAndSaveToStream : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:LoadAndSaveToStream 
            //ExStart:OpeningFromStream
            // Open the stream
            // Read only access is enough for Aspose.Words to load a document
            Stream stream = File.OpenRead(QuickStartDir + "Document.doc");

            Document doc = new Document(stream);
            // You can close the stream now, it is no longer needed because the document is in memory
            stream.Close();
            //ExEnd:OpeningFromStream 

            // ... do something with the document

            // Convert the document to a different format and save to stream
            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Rtf);

            // Rewind the stream position back to zero so it is ready for the next reader
            dstStream.Position = 0;
            //ExEnd:LoadAndSaveToStream 
            // Save the document from stream, to disk
            // Normally you would do something with the stream directly, for example writing the data to a database
            File.WriteAllBytes(ArtifactsDir + "LoadAndSaveToStream.rtf", dstStream.ToArray());
        }
    }
}