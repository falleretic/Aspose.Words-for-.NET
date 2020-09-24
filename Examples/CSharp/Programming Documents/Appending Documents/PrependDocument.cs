using System.Collections;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class PrependDocument : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            Document srcDoc = new Document(JoiningAppendingDir + "Document source.docx");
            Document dstDoc = new Document(JoiningAppendingDir + "Northwind traders.docx");

            // Append the source document to the destination document. This causes the result to have line spacing problems
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

            // Instead prepend the content of the destination document to the start of the source document
            // This results in the same joined document but with no line spacing issues
            DoPrepend(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);

            dstDoc.Save(ArtifactsDir + "PrependDocument.docx");
        }

        public static void DoPrepend(Document dstDoc, Document srcDoc, ImportFormatMode mode)
        {
            // Loop through all sections in the source document
            // Section nodes are immediate children of the Document node so we can just enumerate the Document
            ArrayList sections = new ArrayList(srcDoc.Sections.ToArray());

            // Reverse the order of the sections so they are prepended to start of the destination document in the correct order
            sections.Reverse();

            foreach (Section srcSection in sections)
            {
                // Import the nodes from the source document
                Node dstSection = dstDoc.ImportNode(srcSection, true, mode);

                // Now the new section node can be prepended to the destination document
                // Note how PrependChild is used instead of AppendChild. This is the only line changed compared 
                // To the original method
                dstDoc.PrependChild(dstSection);
            }
        }
    }
}