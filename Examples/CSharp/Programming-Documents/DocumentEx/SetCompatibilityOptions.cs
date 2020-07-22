using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.DocumentEx
{
    class SetCompatibilityOptions : TestDataHelper
    {
        [Test]
        public static void OptimizeFor()
        {
            //ExStart:OptimizeFor
            Document doc = new Document(DocumentDir + "Document.docx");
            doc.CompatibilityOptions.OptimizeFor(Settings.MsWordVersion.Word2016);

            doc.Save(ArtifactsDir + "TestFile.docx");
            //ExEnd:OptimizeFor
        }
    }
}