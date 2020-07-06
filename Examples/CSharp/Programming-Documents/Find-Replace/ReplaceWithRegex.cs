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
        
        public static void RecognizeAndSubstitutionsWithinReplacementPatterns(string dataDir)
        {
            // ExStart:RecognizeAndSubstitutionsWithinReplacementPatterns
            // Create new document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write some text.
            builder.Write("Jason give money to Paul.");

            Regex regex = new Regex(@"([A-z]+) give money to ([A-z]+)");

            // Replace text using substitutions.
            FindReplaceOptions options = new FindReplaceOptions();
            options.UseSubstitutions = true;
            doc.Range.Replace(regex, @"$2 take money from $1", options);
            // ExEnd:RecognizeAndSubstitutionsWithinReplacementPatterns
            Console.WriteLine(doc.GetText()); // The output is: Paul take money from Jason.\f
        }
    }    
}