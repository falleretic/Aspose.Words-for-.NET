using Aspose.Words.Fields;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Fields
{
    class InsertASKFieldWithOutDocumentBuilder : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:InsertASKFieldWithOutDocumentBuilder
            Document doc = new Document(FieldsDir + "in.doc");
            // Get paragraph you want to append this Ask field to
            Paragraph para = (Paragraph) doc.GetChildNodes(NodeType.Paragraph, true)[1];

            // We want to insert an Ask field like this:
            // { ASK \"Test 1\" Test2 \\d Test3 \\o }

            // Create instance of FieldAsk class and lets build the above field code
            FieldAsk field = (FieldAsk) para.AppendField(FieldType.FieldAsk, false);

            // { ASK \"Test 1\" " }
            field.BookmarkName = "Test 1";

            // { ASK \"Test 1\" Test2 }
            field.PromptText = "Test2";

            // { ASK \"Test 1\" Test2 \\d Test3 }
            field.DefaultResponse = "Test3";

            // { ASK \"Test 1\" Test2 \\d Test3 \\o }
            field.PromptOnceOnMailMerge = true;

            // Finally update this Ask field
            field.Update();

            doc.Save(ArtifactsDir + "InsertASKFieldWithOutDocumentBuilder.doc");
            //ExEnd:InsertASKFieldWithOutDocumentBuilder
        }
    }
}