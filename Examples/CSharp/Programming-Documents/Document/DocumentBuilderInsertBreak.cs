namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class DocumentBuilderInsertBreak : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:DocumentBuilderInsertBreak
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("This is page 1.");
            builder.InsertBreak(BreakType.PageBreak);

            builder.Writeln("This is page 2.");
            builder.InsertBreak(BreakType.PageBreak);

            builder.Writeln("This is page 3.");
            doc.Save(ArtifactsDir + "DocumentBuilderInsertBreak.doc");
            //ExEnd:DocumentBuilderInsertBreak
        }
    }
}