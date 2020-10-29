using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace SiteExamples.File_Formats_and_Conversions.Save_Options
{
    internal class WorkingWithDocSaveOptions : SiteExamplesBase
    {
        [Test, Description("Shows how to encrypt document with password.")]
        public void EncryptDocumentWithPassword()
        {
            //ExStart:EncryptDocumentWithPassword
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.Write("Hello world!");
            
            DocSaveOptions saveOptions = new DocSaveOptions();
            saveOptions.Password = "password";
            
            doc.Save(ArtifactsDir + "DocSaveOptions.EncryptDocumentWithPassword.docx", saveOptions);
            //ExEnd:EncryptDocumentWithPassword
        }

        [Test, Description("Shows how to specify not to compress small metafiles.")]
        public void DoNotCompressSmallMetafiles()
        {
            //ExStart:DoNotCompressSmallMetafiles
            Document doc = new Document(MyDir + "Microsoft equation object.docx");
            
            DocSaveOptions saveOptions = new DocSaveOptions();
            saveOptions.AlwaysCompressMetafiles = false;
            
            doc.Save(ArtifactsDir + "DocSaveOptions.NotCompressSmallMetafiles.docx", saveOptions);
            //ExEnd:DoNotCompressSmallMetafiles
        }

        [Test, Description("Shows how to don't save PictureBullet data.")]
        public void DoNotSavePictureBullet()
        {
            //ExStart:DoNotSavePictureBullet
            Document doc = new Document(MyDir + "Image bullet points.docx");
            
            DocSaveOptions saveOptions = (DocSaveOptions)SaveOptions.CreateSaveOptions(SaveFormat.Doc);
            saveOptions.SavePictureBullet = false;
            
            doc.Save(ArtifactsDir + "DocSaveOptions.DoNotSavePictureBullet.docx", saveOptions);
            //ExEnd:DoNotSavePictureBullet
        }
    }
}