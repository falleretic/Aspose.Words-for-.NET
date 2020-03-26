using Aspose.Words.Markup;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class CheckBoxTypeContentControl : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:CheckBoxTypeContentControl
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            StructuredDocumentTag sdtCheckBox = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline);
            // Insert content control into the document
            builder.InsertNode(sdtCheckBox);
            
            doc.Save(ArtifactsDir + "CheckBoxTypeContentControl.docx", SaveFormat.Docx);
            //ExEnd:CheckBoxTypeContentControl
        }
    }
}