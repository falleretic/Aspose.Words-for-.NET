using System.Data;
using Aspose.Words.Fields;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class MailMergeAndConditionalField : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:MailMergeAndConditionalField
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a MERGEFIELD nested inside an IF field
            // Since the statement of the IF field is false, the result of the inner MERGEFIELD will not be displayed
            // and the MERGEFIELD will not receive any data during a mail merge
            FieldIf fieldIf = (FieldIf)builder.InsertField(" IF 1 = 2 ");
            builder.MoveTo(fieldIf.Separator);
            builder.InsertField(" MERGEFIELD  FullName ");

            // We can still count MERGEFIELDs inside false-statement IF fields if we set this flag to true
            doc.MailMerge.UnconditionalMergeFieldsAndRegions = true;

            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("FullName");
            dataTable.Rows.Add("James Bond");

            // Execute the mail merge
            doc.MailMerge.Execute(dataTable);

            // The result will not be visible in the document because the IF field is false, but the inner MERGEFIELD did indeed receive data
            doc.Save(ArtifactsDir + "MailMerge.UnconditionalMergeFieldsAndRegions.docx");
            //ExEnd:MailMergeAndConditionalField
        }
    }
}