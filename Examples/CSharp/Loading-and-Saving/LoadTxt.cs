namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class LoadTxt : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:LoadTxt
            // The encoding of the text file is automatically detected
            Document doc = new Document(LoadingSavingDir + "LoadTxt.txt");
            doc.Save(ArtifactsDir + "LoadTxt.docx");
            //ExEnd:LoadTxt
        }
    }
}