using Aspose.Words.Fields;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Fields
{
    class InsertAuthorField : TestDataHelper
    {
        public static void Run()
        {
            // ExStart:InsertAuthorField
            Document doc = new Document(FieldsDir + "in.doc");
            // Get paragraph you want to append this AUTHOR field to
            Paragraph para = (Paragraph) doc.GetChildNodes(NodeType.Paragraph, true)[1];

            // We want to insert an AUTHOR field like this:
            // { AUTHOR Test1 }

            // Create instance of FieldAuthor class and lets build the above field code
            FieldAuthor field = (FieldAuthor) para.AppendField(FieldType.FieldAuthor, false);

            // { AUTHOR Test1 }
            field.AuthorName = "Test1";

            // Finally update this AUTHOR field
            field.Update();

            doc.Save(ArtifactsDir + "InsertAuthorField.doc");
            //ExEnd:InsertAuthorField
        }
    }
}