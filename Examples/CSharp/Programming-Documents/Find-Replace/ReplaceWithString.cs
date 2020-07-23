using Aspose.Words.Replacing;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class ReplaceWithString : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:ReplaceWithString
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.Writeln("sad mad bad");

            doc.Range.Replace("sad", "bad", new FindReplaceOptions(FindReplaceDirection.Forward));

            doc.Save(ArtifactsDir + "ReplaceWithString.doc");
            //ExEnd:ReplaceWithString
        }
    }
}