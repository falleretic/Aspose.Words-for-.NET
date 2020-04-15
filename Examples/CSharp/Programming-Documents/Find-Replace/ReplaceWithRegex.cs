using System.Text.RegularExpressions;
using Aspose.Words.Replacing;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Find_and_Replace
{
    class ReplaceWithRegex : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:ReplaceWithRegex
            Document doc = new Document(FindReplaceDir + "Document.doc");

            FindReplaceOptions options = new FindReplaceOptions();

            doc.Range.Replace(new Regex("[s|m]ad"), "bad", options);

            doc.Save(ArtifactsDir + "ReplaceWithRegex.doc");
            //ExEnd:ReplaceWithRegex
        }
    }
}