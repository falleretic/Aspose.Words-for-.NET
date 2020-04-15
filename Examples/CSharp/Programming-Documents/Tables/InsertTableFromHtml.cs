using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Tables
{
    class InsertTableFromHtml : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:InsertTableFromHtml
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert the table from HTML. Note that AutoFitSettings does not apply to tables
            // inserted from HTML
            builder.InsertHtml("<table>" +
                               "<tr>" +
                               "<td>Row 1, Cell 1</td>" +
                               "<td>Row 1, Cell 2</td>" +
                               "</tr>" +
                               "<tr>" +
                               "<td>Row 2, Cell 2</td>" +
                               "<td>Row 2, Cell 2</td>" +
                               "</tr>" +
                               "</table>");

            doc.Save(ArtifactsDir + "InsertTableFromHtml.docx");
            //ExEnd:InsertTableFromHtml
        }
    }
}