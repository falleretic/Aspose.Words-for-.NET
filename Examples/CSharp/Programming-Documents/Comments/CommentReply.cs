using System;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class CommentReply : TestDataHelper
    {
        [Test]
        public static void AddRemoveCommentReply()
        {
            //ExStart:AddRemoveCommentReply
            Document doc = new Document(CommentsDir + "TestFile.doc");

            Comment comment = (Comment) doc.GetChild(NodeType.Comment, 0, true);
            // Remove the reply
            comment.RemoveReply(comment.Replies[0]);
            // Add a reply to comment
            comment.AddReply("John Doe", "JD", new DateTime(2017, 9, 25, 12, 15, 0), "New reply");

            doc.Save(ArtifactsDir + "TestFile.doc");
            //ExEnd:AddRemoveCommentReply
        }
    }
}