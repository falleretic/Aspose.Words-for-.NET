using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.DocumentEx
{
    class DocumentBuilderInsertTOC : TestDataHelper 
    {
        [Test]
        public static void Run()
        {
            //ExStart:DocumentBuilderInsertTOC
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a table of contents at the beginning of the document
            builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

            // The newly inserted table of contents will be initially empty
            // It needs to be populated by updating the fields in the document
            //ExStart:UpdateFields
            doc.UpdateFields();
            //ExEnd:UpdateFields
            
            doc.Save(ArtifactsDir + "DocumentBuilderInsertTOC.doc");
            //ExEnd:DocumentBuilderInsertTOC
        }
    }
}