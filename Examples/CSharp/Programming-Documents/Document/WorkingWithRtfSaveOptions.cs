using Aspose.Words.Saving;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class WorkingWithRtfSaveOptions : TestDataHelper
    {
        public static void Run()
        {
            SavingImagesAsWmf();
        }

        public static void SavingImagesAsWmf()
        {
            //ExStart:SavingImagesAsWmf
            Document doc = new Document(DocumentDir + "TestFile.doc");

            RtfSaveOptions saveOpts = new RtfSaveOptions();
            saveOpts.SaveImagesAsWmf = true;

            doc.Save(ArtifactsDir + "output.rtf", saveOpts);
            //ExEnd:SavingImagesAsWmf
        }
    }
}