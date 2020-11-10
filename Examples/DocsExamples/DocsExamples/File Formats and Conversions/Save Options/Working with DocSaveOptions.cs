using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace DocsExamples.File_Formats_and_Conversions.Save_Options
{
    internal class WorkingWithDocSaveOptions : DocsExamplesBase
    {
        [Test]
        public void EncryptDocumentWithPassword()
        {
            //ExStart:EncryptDocumentWithPassword
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.Write("Hello world!");

            DocSaveOptions docSaveOptions = new DocSaveOptions { Password = "password" };

            doc.Save(ArtifactsDir + "WorkingWithDocSaveOptions.EncryptDocumentWithPassword.docx", docSaveOptions);
            //ExEnd:EncryptDocumentWithPassword
        }

        [Test]
        public void DoNotCompressSmallMetafiles()
        {
            //ExStart:DoNotCompressSmallMetafiles
            Document doc = new Document(MyDir + "Microsoft equation object.docx");

            DocSaveOptions docSaveOptions = new DocSaveOptions { AlwaysCompressMetafiles = false };

            doc.Save(ArtifactsDir + "WorkingWithDocSaveOptions.NotCompressSmallMetafiles.docx", docSaveOptions);
            //ExEnd:DoNotCompressSmallMetafiles
        }

        [Test]
        public void DoNotSavePictureBullet()
        {
            //ExStart:DoNotSavePictureBullet
            Document doc = new Document(MyDir + "Image bullet points.docx");

            DocSaveOptions docSaveOptions = new DocSaveOptions { SavePictureBullet = false };

            doc.Save(ArtifactsDir + "WorkingWithDocSaveOptions.DoNotSavePictureBullet.docx", docSaveOptions);
            //ExEnd:DoNotSavePictureBullet
        }
    }
}