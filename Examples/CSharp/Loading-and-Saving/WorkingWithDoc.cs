using Aspose.Words.Saving;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class WorkingWithDoc : TestDataHelper
    {
        public static void Run()
        {
            EncryptDocumentWithPassword();
            AlwaysCompressMetafiles();
            SavePictureBullet();
        }

        public static void EncryptDocumentWithPassword()
        {
            //ExStart:EncryptDocumentWithPassword
            Document doc = new Document(LoadingSavingDir + "Document.doc");
            
            DocSaveOptions docSaveOptions = new DocSaveOptions();
            docSaveOptions.Password = "password";
            
            doc.Save(ArtifactsDir + "EncryptDocumentWithPassword.docx", docSaveOptions);
            //ExEnd:EncryptDocumentWithPassword
        }

        public static void AlwaysCompressMetafiles()
        {
            //ExStart:AlwaysCompressMetafiles
            Document doc = new Document(LoadingSavingDir + "Document.doc");
            
            DocSaveOptions saveOptions = new DocSaveOptions();
            saveOptions.AlwaysCompressMetafiles = false;
            
            doc.Save(ArtifactsDir + "AlwaysCompressMetafiles.doc", saveOptions);
            //ExEnd:AlwaysCompressMetafiles
        }

        public static void SavePictureBullet()
        {
            //ExStart:SavePictureBullet
            Document doc = new Document(LoadingSavingDir + "in.doc");
            
            DocSaveOptions saveOptions = (DocSaveOptions) SaveOptions.CreateSaveOptions(SaveFormat.Doc);
            saveOptions.SavePictureBullet = false;
            
            doc.Save(ArtifactsDir + "SavePictureBullet.doc", saveOptions);
            //ExEnd:SavePictureBullet
        }
    }
}