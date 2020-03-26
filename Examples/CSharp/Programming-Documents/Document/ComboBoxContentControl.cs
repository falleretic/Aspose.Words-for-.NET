using Aspose.Words.Markup;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class ComboBoxContentControl : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:ComboBoxContentControl
            Document doc = new Document();
            StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.ComboBox, MarkupLevel.Block);

            sdt.ListItems.Add(new SdtListItem("Choose an item", "-1"));
            sdt.ListItems.Add(new SdtListItem("Item 1", "1"));
            sdt.ListItems.Add(new SdtListItem("Item 2", "2"));
            doc.FirstSection.Body.AppendChild(sdt);

            doc.Save(ArtifactsDir + "ComboBoxContentControl.docx");
            //ExEnd:ComboBoxContentControl
        }
    }
}