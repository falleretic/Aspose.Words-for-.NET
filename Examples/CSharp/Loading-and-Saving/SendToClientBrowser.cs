namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class SendToClientBrowser : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:SendToClientBrowser
            Document doc = new Document(LoadingSavingDir + "Document.doc");

            // If this method overload is causing a compiler error then you are using the Client Profile DLL whereas 
            // The Aspose.Words .NET 2.0 DLL must be used instead
            doc.Save(ArtifactsDir + "SendToClientBrowser.doc");
            //ExEnd:SendToClientBrowser
        }
    }
}