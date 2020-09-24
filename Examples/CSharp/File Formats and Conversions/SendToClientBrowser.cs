using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class SendToClientBrowser : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:SendToClientBrowser
            Document doc = new Document(LoadingSavingDir + "Document.docx");

            // If this method overload is causing a compiler error then you are using the Client Profile DLL whereas 
            // The Aspose.Words .NET 2.0 DLL must be used instead
            doc.Save(ArtifactsDir + "SendToClientBrowser.doc");
            //ExEnd:SendToClientBrowser
        }
    }
}