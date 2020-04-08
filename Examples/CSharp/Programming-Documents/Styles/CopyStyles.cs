namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Styles
{
    class CopyStyles : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:CopyStylesFromDocument
            Document doc = new Document(StyleDir + "template.docx");
            Document target = new Document(StyleDir + "TestFile.doc");
            
            target.CopyStylesFromTemplate(doc);
            
            doc.Save(ArtifactsDir + "CopyStyles.docx");
            //ExEnd:CopyStylesFromDocument
        }
    }
}