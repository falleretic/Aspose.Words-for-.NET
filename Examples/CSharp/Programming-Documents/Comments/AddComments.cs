using System;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Comments
{
    class AddComments : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:AddComments
            // ExStart:CreateSimpleDocumentUsingDocumentBuilder
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Write("Some text is added.");
            //ExEnd:CreateSimpleDocumentUsingDocumentBuilder
            
            Comment comment = new Comment(doc, "Awais Hafeez", "AH", DateTime.Today);
            builder.CurrentParagraph.AppendChild(comment);
            comment.Paragraphs.Add(new Paragraph(doc));
            comment.FirstParagraph.Runs.Add(new Run(doc, "Comment text."));

            doc.Save(ArtifactsDir + "Comments.doc");
            //ExEnd:AddComments

            Console.WriteLine("\nComments added successfully.");
        }
    }
}