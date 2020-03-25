using System;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class CommentReply : TestDataHelper
    {
        public static void Run()
        {
            AddRemoveCommentReply();
        }

        static void AddRemoveCommentReply()
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
            
            Console.WriteLine("\nComment's reply is removed successfully.");
        }
    }
}