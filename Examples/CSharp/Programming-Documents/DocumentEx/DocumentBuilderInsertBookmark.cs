using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.DocumentEx
{
    class DocumentBuilderInsertBookmark : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:DocumentBuilderInsertBookmark
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.StartBookmark("FineBookmark");
            builder.Writeln("This is just a fine bookmark.");
            builder.EndBookmark("FineBookmark");

            doc.Save(ArtifactsDir + "DocumentBuilderInsertBookmark.doc");
            //ExEnd:DocumentBuilderInsertBookmark
        }
    }
}