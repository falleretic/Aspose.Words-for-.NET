using System.Collections;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.DocumentEx
{
    class ExtractContentBetweenCommentRange : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:ExtractContentBetweenCommentRange
            Document doc = new Document(DocumentDir + "Document.docx");

            // This is a quick way of getting both comment nodes
            // Your code should have a proper method of retrieving each corresponding start and end node
            CommentRangeStart commentStart = (CommentRangeStart) doc.GetChild(NodeType.CommentRangeStart, 0, true);
            CommentRangeEnd commentEnd = (CommentRangeEnd) doc.GetChild(NodeType.CommentRangeEnd, 0, true);

            // Firstly extract the content between these nodes including the comment as well
            ArrayList extractedNodesInclusive = Common.ExtractContent(commentStart, commentEnd, true);
            Document dstDoc = Common.GenerateDocument(doc, extractedNodesInclusive);
            dstDoc.Save(ArtifactsDir + "TestFile.CommentInclusive.doc");

            // Secondly extract the content between these nodes without the comment
            ArrayList extractedNodesExclusive = Common.ExtractContent(commentStart, commentEnd, false);
            dstDoc = Common.GenerateDocument(doc, extractedNodesExclusive);
            dstDoc.Save(ArtifactsDir + "TestFile.CommentExclusive.doc");
            //ExEnd:ExtractContentBetweenCommentRange
        }
    }
}