using System.Collections;
using Aspose.Words.Fields;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.DocumentEx
{
    class ExtractContentUsingField : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:ExtractContentUsingField
            Document doc = new Document(DocumentDir + "TestFile.doc");

            // Use a document builder to retrieve the field start of a merge field
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Pass the first boolean parameter to get the DocumentBuilder to move to the FieldStart of the field
            // We could also get FieldStarts of a field using GetChildNode method as in the other examples
            builder.MoveToMergeField("Fullname", false, false);

            // The builder cursor should be positioned at the start of the field
            FieldStart startField = (FieldStart) builder.CurrentNode;
            Paragraph endPara = (Paragraph) doc.FirstSection.GetChild(NodeType.Paragraph, 5, true);

            // Extract the content between these nodes in the document
            // Don't include these markers in the extraction
            ArrayList extractedNodes = Common.ExtractContent(startField, endPara, false);

            // Insert the content into a new separate document and save it to disk
            Document dstDoc = Common.GenerateDocument(doc, extractedNodes);
            dstDoc.Save(ArtifactsDir + "TestFile.doc");
            //ExEnd:ExtractContentUsingField
        }
    }
}