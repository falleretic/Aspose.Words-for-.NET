using Aspose.Words.Saving;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.DocumentEx
{
    class WorkingWithRtfSaveOptions : TestDataHelper
    {
        [Test]
        public static void SavingImagesAsWmf()
        {
            //ExStart:SavingImagesAsWmf
            Document doc = new Document(DocumentDir + "Document.docx");

            RtfSaveOptions saveOpts = new RtfSaveOptions();
            saveOpts.SaveImagesAsWmf = true;

            doc.Save(ArtifactsDir + "output.rtf", saveOpts);
            //ExEnd:SavingImagesAsWmf
        }
    }
}