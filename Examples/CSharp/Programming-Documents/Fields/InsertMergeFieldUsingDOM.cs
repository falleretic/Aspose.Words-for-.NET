using Aspose.Words.Fields;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Fields
{
    class InsertMergeFieldUsingDOM : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:InsertMergeFieldUsingDOM
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Get paragraph you want to append this merge field to
            Paragraph para = (Paragraph) doc.GetChildNodes(NodeType.Paragraph, true)[0];

            // Move cursor to this paragraph
            builder.MoveTo(para);

            // We want to insert a merge field like this:
            // { " MERGEFIELD Test1 \\b Test2 \\f Test3 \\m \\v" }

            // Create instance of FieldMergeField class and lets build the above field code
            FieldMergeField field = (FieldMergeField) builder.InsertField(FieldType.FieldMergeField, false);

            // { " MERGEFIELD Test1" }
            field.FieldName = "Test1";

            // { " MERGEFIELD Test1 \\b Test2" }
            field.TextBefore = "Test2";

            // { " MERGEFIELD Test1 \\b Test2 \\f Test3 }
            field.TextAfter = "Test3";

            // { " MERGEFIELD Test1 \\b Test2 \\f Test3 \\m" }
            field.IsMapped = true;

            // { " MERGEFIELD Test1 \\b Test2 \\f Test3 \\m \\v" }
            field.IsVerticalFormatting = true;

            // Finally update this merge field
            field.Update();

            doc.Save(ArtifactsDir + "InsertMergeFieldUsingDOM.doc");
            //ExEnd:InsertMergeFieldUsingDOM
        }
    }
}