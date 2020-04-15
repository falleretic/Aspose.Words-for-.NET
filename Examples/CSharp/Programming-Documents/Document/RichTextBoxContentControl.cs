using Aspose.Words.Markup;
using System.Drawing;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class RichTextBoxContentControl : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:RichTextBoxContentControl
            Document doc = new Document();
            StructuredDocumentTag sdtRichText = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block);

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