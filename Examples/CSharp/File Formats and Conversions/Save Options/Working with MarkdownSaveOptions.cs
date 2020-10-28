using Aspose.Words.Saving;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.File_Formats_and_Conversions.Save_Options
{
    internal class WorkingWithMarkdownSaveOptions : TestDataHelper
    {
        [Test, Description("Shows how to save the document to Markdown format.")]
        public void SaveToMarkdownDocument()
        {
            //ExStart:SaveToMarkdownDocument
            DocumentBuilder builder = new DocumentBuilder();
            builder.Writeln("Some text!");

            MarkdownSaveOptions saveOptions = (MarkdownSaveOptions)SaveOptions.CreateSaveOptions(SaveFormat.Markdown);
            
            builder.Document.Save(ArtifactsDir + "MarkdownSaveOptions.MarkdownDocument.md", saveOptions);
            //ExEnd:SaveToMarkdownDocument
        }

        [Test, Description("Shows how to specify table content alignment.")]
        public void ExportIntoMarkdownWithTableContentAlignment()
        {
            //ExStart:ExportIntoMarkdownWithTableContentAlignment
            DocumentBuilder builder = new DocumentBuilder();

            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
            builder.Write("Cell1");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.Write("Cell2");

            MarkdownSaveOptions saveOptions = new MarkdownSaveOptions();
            // Makes all paragraphs inside table to be aligned to Left. 
            saveOptions.TableContentAlignment = TableContentAlignment.Left;
            builder.Document.Save(ArtifactsDir + "MarkdownSaveOptions.LeftTableContentAlignment.md", saveOptions);

            // Makes all paragraphs inside table to be aligned to Right. 
            saveOptions.TableContentAlignment = TableContentAlignment.Right;
            builder.Document.Save(ArtifactsDir + "MarkdownSaveOptions.RightTableContentAlignment.md", saveOptions);

            // Makes all paragraphs inside table to be aligned to Center. 
            saveOptions.TableContentAlignment = TableContentAlignment.Center;
            builder.Document.Save(ArtifactsDir + "MarkdownSaveOptions.CenterTableContentAlignment.md", saveOptions);

            // Makes all paragraphs inside table to be aligned automatically.
            // The alignment in this case will be taken from the first paragraph in corresponding table column.
            saveOptions.TableContentAlignment = TableContentAlignment.Auto;
            builder.Document.Save(ArtifactsDir + "MarkdownSaveOptions.AutoTableContentAlignment.md", saveOptions);
            //ExEnd:ExportIntoMarkdownWithTableContentAlignment
        }
    }
}
