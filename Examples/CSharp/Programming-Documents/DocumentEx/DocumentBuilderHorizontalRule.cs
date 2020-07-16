using System.Drawing;
using Aspose.Words.Drawing;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.DocumentEx
{
    class DocumentBuilderHorizontalRule : TestDataHelper
    {
        [Test]
        public static void DocumentBuilderInsertHorizontalRule()
        {
            //ExStart:DocumentBuilderInsertHorizontalRule
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Insert a horizontal rule shape into the document.");
            builder.InsertHorizontalRule();

            doc.Save(ArtifactsDir + "DocumentBuilder.InsertHorizontalRule.doc");
            //ExEnd:DocumentBuilderInsertHorizontalRule
        }

        [Test]
        public static void DocumentBuilderHorizontalRuleFormat()
        {
            //ExStart:DocumentBuilderHorizontalRuleFormat
            DocumentBuilder builder = new DocumentBuilder();

            Shape shape = builder.InsertHorizontalRule();
            HorizontalRuleFormat horizontalRuleFormat = shape.HorizontalRuleFormat;

            horizontalRuleFormat.Alignment = HorizontalRuleAlignment.Center;
            horizontalRuleFormat.WidthPercent = 70;
            horizontalRuleFormat.Height = 3;
            horizontalRuleFormat.Color = Color.Blue;
            horizontalRuleFormat.NoShade = true;

            builder.Document.Save(ArtifactsDir + "HorizontalRuleFormat.docx");
            //ExEnd:DocumentBuilderHorizontalRuleFormat
        }
    }
}