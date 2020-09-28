using Aspose.Words.Saving;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class WorkingWithDoc : TestDataHelper
    {
        [Test]
        public static void EncryptDocumentWithPassword()
        {
            //ExStart:EncryptDocumentWithPassword
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.Write("Hello world!");
            
            DocSaveOptions docSaveOptions = new DocSaveOptions();
            docSaveOptions.Password = "password";
            
            doc.Save(ArtifactsDir + "DocSaveOptions.EncryptDocumentWithPassword.docx", docSaveOptions);
            //ExEnd:EncryptDocumentWithPassword
        }

        [Test]
        public static void AlwaysCompressMetafiles()
        {
            //ExStart:AlwaysCompressMetafiles
            Document doc = new Document(LoadingSavingDir + "Microsoft equation object.docx");
            
            DocSaveOptions saveOptions = new DocSaveOptions();
            saveOptions.AlwaysCompressMetafiles = false;
            
            doc.Save(ArtifactsDir + "DocSaveOptions.AlwaysCompressMetafiles.docx", saveOptions);
            //ExEnd:AlwaysCompressMetafiles
        }

        [Test]
        public static void SavePictureBullet()
        {
            //ExStart:SavePictureBullet
            Document doc = new Document(LoadingSavingDir + "Image bullet points.docx");
            
            DocSaveOptions saveOptions = (DocSaveOptions) SaveOptions.CreateSaveOptions(SaveFormat.Doc);
            saveOptions.SavePictureBullet = false;
            
            doc.Save(ArtifactsDir + "DocSaveOptions.SavePictureBullet.docx", saveOptions);
            //ExEnd:SavePictureBullet
        }
    }
}