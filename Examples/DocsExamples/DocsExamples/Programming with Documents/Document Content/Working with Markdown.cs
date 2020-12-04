using Aspose.Words;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents.Document_Content
{
    class WorkingWithMarkdownFeatures : DocsExamplesBase
    {
        [Test]
        public static void Emphases()
        {
            //ExStart:Emphases
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Markdown treats asterisks (*) and underscores (_) as indicators of emphasis.");
            builder.Write("You can write ");

            builder.Font.Bold = true;
            builder.Write("bold");

            builder.Font.Bold = false;
            builder.Write(" or ");

            builder.Font.Italic = true;
            builder.Write("italic");

            builder.Font.Italic = false;
            builder.Writeln(" text. ");

            builder.Write("You can also write ");
            builder.Font.Bold = true;

            builder.Font.Italic = true;
            builder.Write("BoldItalic");

            builder.Font.Bold = false;
            builder.Font.Italic = false;
            builder.Write("text.");

            builder.Document.Save(ArtifactsDir + "WorkingWithMarkdownFeatures.Emphases.md");
            //ExEnd:Emphases
        }

        [Test]
        public static void Headings()
        {
            //ExStart:Headings
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // By default Heading styles in Word may have bold and italic formatting.
            // If we do not want the text to be emphasized, set these properties explicitly to false.
            builder.Font.Bold = false;
            builder.Font.Italic = false;

            builder.Writeln("The following produces headings:");
            builder.ParagraphFormat.Style = doc.Styles["Heading 1"];
            builder.Writeln("Heading1");
            builder.ParagraphFormat.Style = doc.Styles["Heading 2"];
            builder.Writeln("Heading2");
            builder.ParagraphFormat.Style = doc.Styles["Heading 3"];
            builder.Writeln("Heading3");
            builder.ParagraphFormat.Style = doc.Styles["Heading 4"];
            builder.Writeln("Heading4");
            builder.ParagraphFormat.Style = doc.Styles["Heading 5"];
            builder.Writeln("Heading5");
            builder.ParagraphFormat.Style = doc.Styles["Heading 6"];
            builder.Writeln("Heading6");

            // Note that the emphases are also allowed inside Headings.
            builder.Font.Bold = true;
            builder.ParagraphFormat.Style = doc.Styles["Heading 1"];
            builder.Writeln("Bold Heading1");

            doc.Save(ArtifactsDir + "WorkingWithMarkdownFeatures.Headings.md");
            //ExEnd:Headings
        }

        [Test]
        public static void BlockQuotes()
        {
            //ExStart:BlockQuotes
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("We support blockquotes in Markdown:");
            
            builder.ParagraphFormat.Style = doc.Styles["Quote"];
            builder.Writeln("Lorem");
            builder.Writeln("ipsum");
            
            builder.ParagraphFormat.Style = doc.Styles["Normal"];
            builder.Writeln("The quotes can be of any level and can be nested:");
            
            Style quoteLevel3 = doc.Styles.Add(StyleType.Paragraph, "Quote2");
            builder.ParagraphFormat.Style = quoteLevel3;
            builder.Writeln("Quote level 3");
            
            Style quoteLevel4 = doc.Styles.Add(StyleType.Paragraph, "Quote3");
            builder.ParagraphFormat.Style = quoteLevel4;
            builder.Writeln("Nested quote level 4");
            
            builder.ParagraphFormat.Style = doc.Styles["Quote"];
            builder.Writeln();
            builder.Writeln("Back to first level");
            
            Style quoteLevel1WithHeading = doc.Styles.Add(StyleType.Paragraph, "Quote Heading 3");
            builder.ParagraphFormat.Style = quoteLevel1WithHeading;
            builder.Write("Headings are allowed inside Quotes");

            doc.Save(ArtifactsDir + "WorkingWithMarkdownFeatures.BlockQuotes.md");
            //ExEnd:BlockQuotes
        }

        [Test]
        public static void ReadMarkdownDocument()
        {
            //ExStart:ReadMarkdownDocument
            Document doc = new Document(MyDir + "Quotes.md");

            // Let's remove Heading formatting from a Quote in the very last paragraph.
            Paragraph paragraph = doc.FirstSection.Body.LastParagraph;
            paragraph.ParagraphFormat.Style = doc.Styles["Quote"];

            doc.Save(ArtifactsDir + "WorkingWithMarkdownFeatures.ReadMarkdownDocument.md");
            //ExEnd:ReadMarkdownDocument
        }

        [Test]
        public static void HorizontalRule()
        {
            //ExStart:HorizontalRule
            DocumentBuilder builder = new DocumentBuilder(new Document());

            builder.Writeln("We support Horizontal rules (Thematic breaks) in Markdown:");
            builder.InsertHorizontalRule();

            builder.Document.Save(ArtifactsDir + "WorkingWithMarkdownFeatures.HorizontalRuleExample.md");
            //ExEnd:HorizontalRule
        }
    }
}