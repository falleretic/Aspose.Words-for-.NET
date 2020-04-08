using Aspose.Words.Fields;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Fields
{
    class InsertTOAFieldWithoutDocumentBuilder : TestDataHelper
    {
        public static void Run()
        {
            // ExStart:InsertTOAFieldWithoutDocumentBuilder
            Document doc = new Document(FieldsDir + "in.doc");
            // Get paragraph you want to append this TOA field to
            Paragraph para = (Paragraph) doc.GetChildNodes(NodeType.Paragraph, true)[1];

            // We want to insert TA and TOA fields like this:
            // { TA  \c 1 \l "Value 0" }
            // { TOA  \c 1 }

            // Create instance of FieldAsk class and lets build the above field code
            FieldTA fieldTA = (FieldTA) para.AppendField(FieldType.FieldTOAEntry, false);
            fieldTA.EntryCategory = "1";
            fieldTA.LongCitation = "Value 0";

            doc.FirstSection.Body.AppendChild(para);

            para = new Paragraph(doc);

            // Create instance of FieldToa class
            FieldToa fieldToa = (FieldToa) para.AppendField(FieldType.FieldTOA, false);
            fieldToa.EntryCategory = "1";
            doc.FirstSection.Body.AppendChild(para);

            // Finally update this TOA field
            fieldToa.Update();

            doc.Save(ArtifactsDir + "InsertTOAFieldWithoutDocumentBuilder.doc");
            //ExEnd:InsertTOAFieldWithoutDocumentBuilder
        }
    }
}