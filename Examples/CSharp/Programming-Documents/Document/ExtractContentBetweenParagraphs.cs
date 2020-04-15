using System.Collections;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class ExtractContentBetweenParagraphs : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:ExtractContentBetweenParagraphs
            Document doc = new Document(DocumentDir + "TestFile.doc");

            // Gather the nodes
            // The GetChild method uses 0-based index
            Paragraph startPara = (Paragraph) doc.FirstSection.Body.GetChild(NodeType.Paragraph, 6, true);
            Paragraph endPara = (Paragraph) doc.FirstSection.Body.GetChild(NodeType.Paragraph, 10, true);
            // Extract the content between these nodes in the document
            // Include these markers in the extraction
            ArrayList extractedNodes = Common.ExtractContent(startPara, endPara, true);

            // Insert the content into a new separate document and save it to disk
            Document dstDoc = Common.GenerateDocument(doc, extractedNodes);
            dstDoc.Save(ArtifactsDir + "TestFile.doc");
            //ExEnd:ExtractContentBetweenParagraphs
        }
    }
}