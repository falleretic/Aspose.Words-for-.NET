﻿using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.DocumentEx
{
    class RemoveBreaks : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            // ExStart:OpenFromFile
            Document doc = new Document(DocumentDir + "Document.docx");
            // ExEnd:OpenFromFile

            // Remove the page and section breaks from the document
            // In Aspose.Words section breaks are represented as separate Section nodes in the document
            // To remove these separate sections the sections are combined
            RemovePageBreaks(doc);
            RemoveSectionBreaks(doc);

            doc.Save(ArtifactsDir + "TestFile.doc");
        }

        //ExStart:RemovePageBreaks
        private static void RemovePageBreaks(Document doc)
        {
            // Retrieve all paragraphs in the document
            NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

            // Iterate through all paragraphs
            foreach (Paragraph para in paragraphs)
            {
                // If the paragraph has a page break before set then clear it
                if (para.ParagraphFormat.PageBreakBefore)
                    para.ParagraphFormat.PageBreakBefore = false;

                // Check all runs in the paragraph for page breaks and remove them
                foreach (Run run in para.Runs)
                {
                    if (run.Text.Contains(ControlChar.PageBreak))
                        run.Text = run.Text.Replace(ControlChar.PageBreak, string.Empty);
                }
            }
        }
        //ExEnd:RemovePageBreaks

        //ExStart:RemoveSectionBreaks
        private static void RemoveSectionBreaks(Document doc)
        {
            // Loop through all sections starting from the section that precedes the last one 
            // And moving to the first section
            for (int i = doc.Sections.Count - 2; i >= 0; i--)
            {
                // Copy the content of the current section to the beginning of the last section
                doc.LastSection.PrependContent(doc.Sections[i]);
                // Remove the copied section
                doc.Sections[i].Remove();
            }
        }
        //ExEnd:RemoveSectionBreaks
    }
}