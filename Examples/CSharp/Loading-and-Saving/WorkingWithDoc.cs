using Aspose.Words.Saving;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class WorkingWithDoc : TestDataHelper
    {
        [Test]
        public static void EncryptDocumentWithPassword()
        {
            //ExStart:EncryptDocumentWithPassword
            Document doc = new Document(LoadingSavingDir + "Document.doc");
            
            DocSaveOptions docSaveOptions = new DocSaveOptions();
            docSaveOptions.Password = "password";
            
            doc.Save(ArtifactsDir + "EncryptDocumentWithPassword.docx", docSaveOptions);
            //ExEnd:EncryptDocumentWithPassword
        }

        [Test]
        public static void AlwaysCompressMetafiles()
        {
            //ExStart:AlwaysCompressMetafiles
            Document doc = new Document(LoadingSavingDir + "Document.doc");
            
            DocSaveOptions saveOptions = new DocSaveOptions();
            saveOptions.AlwaysCompressMetafiles = false;
            
            doc.Save(ArtifactsDir + "AlwaysCompressMetafiles.doc", saveOptions);
            //ExEnd:AlwaysCompressMetafiles
        }

        [Test]
        public static void SavePictureBullet()
        {
            //ExStart:SavePictureBullet
            Document doc = new Document(LoadingSavingDir + "Document.doc");
            
            DocSaveOptions saveOptions = (DocSaveOptions) SaveOptions.CreateSaveOptions(SaveFormat.Doc);
            saveOptions.SavePictureBullet = false;
            
            doc.Save(ArtifactsDir + "SavePictureBullet.doc", saveOptions);
            //ExEnd:SavePictureBullet
        }
    }
}