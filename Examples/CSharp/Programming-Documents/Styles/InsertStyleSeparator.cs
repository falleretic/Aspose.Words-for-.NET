using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Styles
{
    class InsertStyleSeparator : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:ParagraphInsertStyleSeparator
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Style paraStyle = builder.Document.Styles.Add(StyleType.Paragraph, "MyParaStyle");
            paraStyle.Font.Bold = false;
            paraStyle.Font.Size = 8;
            paraStyle.Font.Name = "Arial";

            // Append text with "Heading 1" style
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Write("Heading 1");
            builder.InsertStyleSeparator();

            // Append text with another style
            builder.ParagraphFormat.StyleName = paraStyle.Name;
            builder.Write("This is text with some other formatting ");

            doc.Save(ArtifactsDir + "InsertStyleSeparator.docx");
            //ExEnd:ParagraphInsertStyleSeparator
        }
    }
}