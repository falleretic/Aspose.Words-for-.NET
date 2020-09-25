using System;
using Aspose.Words.Layout;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.DocumentEx
{
    class WorkingWithRevisions : TestDataHelper
    {
        [Test]
        public static void AcceptRevisions()
        {
            //ExStart:AcceptAllRevisions
            Document doc = new Document(DocumentDir + "Document.docx");

            // Start tracking and make some revisions
            doc.StartTrackRevisions("Author");
            doc.FirstSection.Body.AppendParagraph("Hello world!");

            // Revisions will now show up as normal text in the output document
            doc.AcceptAllRevisions();
            
            doc.Save(ArtifactsDir + "Document.AcceptedRevisions.doc");
            //ExEnd:AcceptAllRevisions
        }

        [Test]
        public static void GetRevisionTypes()
        {
            //ExStart:GetRevisionTypes
            Document doc = new Document(DocumentDir + "Revisions.docx");

            ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;
            for (int i = 0; i < paragraphs.Count; i++)
            {
                if (paragraphs[i].IsMoveFromRevision)
                    Console.WriteLine("The paragraph {0} has been moved (deleted).", i);
                if (paragraphs[i].IsMoveToRevision)
                    Console.WriteLine("The paragraph {0} has been moved (inserted).", i);
            }
            //ExEnd:GetRevisionTypes
        }

        [Test]
        public static void GetRevisionGroups()
        {
            //ExStart:GetRevisionGroups
            Document doc = new Document(DocumentDir + "Revisions.docx");

            foreach (RevisionGroup group in doc.Revisions.Groups)
            {
                Console.WriteLine("{0}, {1}:", group.Author, group.RevisionType);
                Console.WriteLine(group.Text);
            }
            //ExEnd:GetRevisionGroups
        }

        [Test]
        public static void SetShowCommentsInPDF()
        {
            //ExStart:SetShowCommentsinPDF
            Document doc = new Document(DocumentDir + "Revisions.docx");

            // Do not render the comments in PDF
            doc.LayoutOptions.ShowComments = false;
            doc.Save(ArtifactsDir + "RemoveCommentsInPDF.pdf");
            //ExEnd:SetShowCommentsinPDF
        }

        [Test]
        public static void SetShowInBalloons()
        {
            //ExStart:SetShowInBalloons
            Document doc = new Document(DocumentDir + "Revisions.docx");

            // Renders insert and delete revisions inline, format revisions in balloons
            doc.LayoutOptions.RevisionOptions.ShowInBalloons = ShowInBalloons.Format;

            // Renders insert revisions inline, delete and format revisions in balloons
            //doc.LayoutOptions.RevisionOptions.ShowInBalloons = ShowInBalloons.FormatAndDelete;

            doc.Save(ArtifactsDir + "SetShowInBalloons.pdf");
            //ExEnd:SetShowInBalloons
        }

        [Test]
        public static void GetRevisionGroupDetails()
        {
            //ExStart:GetRevisionGroupDetails
            Document doc = new Document(DocumentDir + "Revisions.docx");

            foreach (Revision revision in doc.Revisions)
            {
                string groupText = revision.Group != null
                    ? "Revision group text: " + revision.Group.Text
                    : "Revision has no group";

                Console.WriteLine("Type: " + revision.RevisionType);
                Console.WriteLine("Author: " + revision.Author);
                Console.WriteLine("Date: " + revision.DateTime);
                Console.WriteLine("Revision text: " + revision.ParentNode.ToString(SaveFormat.Text));
                Console.WriteLine(groupText);
            }
            //ExEnd:GetRevisionGroupDetails
        }

        [Test]
        public static void AccessRevisedVersion()
        {
            //ExStart:AccessRevisedVersion
            Document doc = new Document(DocumentDir + "Revisions.docx");
            doc.UpdateListLabels();

            // Switch to the revised version of the document
            doc.RevisionsView = RevisionsView.Final;

            foreach (Revision revision in doc.Revisions)
            {
                if (revision.ParentNode.NodeType == NodeType.Paragraph)
                {
                    Paragraph paragraph = (Paragraph) revision.ParentNode;
                    if (paragraph.IsListItem)
                    {
                        // Print revised version of LabelString and ListLevel
                        Console.WriteLine(paragraph.ListLabel.LabelString);
                        Console.WriteLine(paragraph.ListFormat.ListLevel);
                    }
                }
            }
            //ExEnd:AccessRevisedVersion
        }
    }
}