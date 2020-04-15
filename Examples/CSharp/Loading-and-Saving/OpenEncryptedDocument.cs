using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class OpenEncryptedDocument : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:OpenEncryptedDocument
            Document doc = new Document(LoadingSavingDir + "LoadEncrypted.docx", new LoadOptions("aspose"));
            //ExEnd:OpenEncryptedDocument
        }
    }
}