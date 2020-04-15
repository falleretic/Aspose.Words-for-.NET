using Aspose.Words.Replacing;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Find_and_Replace
{
    class ReplaceHtmlTextWithMeta_Characters : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:ReplaceHtmlTextWithMetaCharacters
            Document doc = new Document();

            // Use a document builder to add content to the document
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Write("{PLACEHOLDER}");

            FindReplaceOptions findReplaceOptions = new FindReplaceOptions
            {
                ReplacingCallback = new FindAndInsertHtml(),
                PreserveMetaCharacters = true
            };

            doc.Range.Replace("{PLACEHOLDER}", "<p>&ldquo;Some Text&rdquo;</p>", findReplaceOptions);

            doc.Save(ArtifactsDir + "ReplaceHtmlTextWithMetaCharacters.doc");
            //ExEnd:ReplaceHtmlTextWithMetaCharacters
        }
    }

    //ExStart:ReplaceHtmlFindAndInsertHtml
    public sealed class FindAndInsertHtml : IReplacingCallback
    {
        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
        {
            // This is a Run node that contains either the beginning or the complete match
            Node currentNode = e.MatchNode;

            // Create document builder and insert MergeField
            DocumentBuilder builder = new DocumentBuilder(e.MatchNode.Document as Document);
            builder.MoveTo(currentNode);
            builder.InsertHtml(e.Replacement);

            currentNode.Remove();

            // Signal to the replace engine to do nothing because we have already done all what we wanted
            return ReplaceAction.Skip;
        }
    }
    //ExEnd:ReplaceHtmlFindAndInsertHtml
}