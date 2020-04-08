using Aspose.Words.Replacing;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Find_and_Replace
{
    class ReplaceWithString : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:ReplaceWithString
            Document doc = new Document(FindReplaceDir + "Document.doc");
            doc.Range.Replace("sad", "bad", new FindReplaceOptions(FindReplaceDirection.Forward));

            doc.Save(ArtifactsDir + "ReplaceWithString.doc");
            //ExEnd:ReplaceWithString
        }
    }
}