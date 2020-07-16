using System.Drawing;
using Aspose.Words.Markup;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.DocumentEx
{
    class RichTextBoxContentControl : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:RichTextBoxContentControl
            Document doc = new Document();
            Markup.StructuredDocumentTag sdtRichText = new Markup.StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block);

            Paragraph para = new Paragraph(doc);
            Run run = new Run(doc);
            run.Text = "Hello World";
            run.Font.Color = Color.Green;
            para.Runs.Add(run);
            sdtRichText.ChildNodes.Add(para);
            doc.FirstSection.Body.AppendChild(sdtRichText);

            doc.Save(ArtifactsDir + "RichTextBoxContentControl.docx");
            //ExEnd:RichTextBoxContentControl
        }
    }
}