using Aspose.Words.Fields;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Fields
{
    class InsertFieldIncludeTextWithoutDocumentBuilder : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:InsertFieldIncludeTextWithoutDocumentBuilder
            Document doc = new Document(FieldsDir + "in.doc");
            // Get paragraph you want to append this INCLUDETEXT field to
            Paragraph para = (Paragraph) doc.GetChildNodes(NodeType.Paragraph, true)[1];

            // We want to insert an INCLUDETEXT field like this:
            // { INCLUDETEXT  "file path" }

            // Create instance of FieldAsk class and lets build the above field code
            FieldIncludeText fieldIncludeText = (FieldIncludeText) para.AppendField(FieldType.FieldIncludeText, false);
            fieldIncludeText.BookmarkName = "bookmark";
            fieldIncludeText.SourceFullName = FieldsDir + "IncludeText.docx";

            doc.FirstSection.Body.AppendChild(para);

            // Finally update this IncludeText field
            fieldIncludeText.Update();

            doc.Save(ArtifactsDir + "InsertIncludeFieldWithoutDocumentBuilder.doc");
            //ExEnd:InsertFieldIncludeTextWithoutDocumentBuilder
        }
    }
}