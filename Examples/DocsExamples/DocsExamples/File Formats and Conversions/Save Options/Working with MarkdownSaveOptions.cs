﻿using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace DocsExamples.File_Formats_and_Conversions.Save_Options
{
    internal class WorkingWithMarkdownSaveOptions : DocsExamplesBase
    {
        [Test]
        public void SaveToMarkdownDocument()
        {
            //ExStart:SaveToMarkdownDocument
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.Writeln("Some text!");

            doc.Save(ArtifactsDir + "WorkingWithMarkdownSaveOptions.MarkdownDocument.md", new MarkdownSaveOptions());
            //ExEnd:SaveToMarkdownDocument
        }

        [Test]
        public void ExportIntoMarkdownWithTableContentAlignment()
        {
            //ExStart:ExportIntoMarkdownWithTableContentAlignment
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
            builder.Write("Cell1");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.Write("Cell2");

            // Makes all paragraphs inside the table to be aligned.
            MarkdownSaveOptions markdownSaveOptions = new MarkdownSaveOptions
            {
                TableContentAlignment = TableContentAlignment.Left
            };
            doc.Save(ArtifactsDir + "WorkingWithMarkdownSaveOptions.LeftTableContentAlignment.md", markdownSaveOptions);

            markdownSaveOptions.TableContentAlignment = TableContentAlignment.Right;
            doc.Save(ArtifactsDir + "WorkingWithMarkdownSaveOptions.RightTableContentAlignment.md", markdownSaveOptions);

            markdownSaveOptions.TableContentAlignment = TableContentAlignment.Center;
            doc.Save(ArtifactsDir + "WorkingWithMarkdownSaveOptions.CenterTableContentAlignment.md", markdownSaveOptions);

            // The alignment in this case will be taken from the first paragraph in corresponding table column.
            markdownSaveOptions.TableContentAlignment = TableContentAlignment.Auto;
            doc.Save(ArtifactsDir + "WorkingWithMarkdownSaveOptions.AutoTableContentAlignment.md", markdownSaveOptions);
            //ExEnd:ExportIntoMarkdownWithTableContentAlignment
        }
    }
}
