using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Joining_and_Appending
{
    internal class AppendDocumentManually : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:AppendDocumentManually
            Document srcDoc = new Document(JoiningAppendingDir + "TestFile.Source.doc");
            Document dstDoc = new Document(JoiningAppendingDir + "TestFile.Destination.doc");
            
            // Loop through all sections in the source document
            // Section nodes are immediate children of the Document node so we can just enumerate the Document
            foreach (Section srcSection in srcDoc)
            {
                // Because we are copying a section from one document to another, 
                // it is required to import the Section node into the destination document
                // This adjusts any document-specific references to styles, lists, etc.
                //
                // Importing a node creates a copy of the original node, but the copy
                // Is ready to be inserted into the destination document
                Node dstSection = dstDoc.ImportNode(srcSection, true, ImportFormatMode.KeepSourceFormatting);

                // Now the new section node can be appended to the destination document
                dstDoc.AppendChild(dstSection);
            }

            dstDoc.Save(ArtifactsDir + "AppendDocumentManually.docx");
            //ExEnd:AppendDocumentManually
        }
    }
}