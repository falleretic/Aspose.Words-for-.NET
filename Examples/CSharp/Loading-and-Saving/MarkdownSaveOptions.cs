using Aspose.Words.Saving;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Loading_and_Saving
{
    class MarkdownSaveOptions : TestDataHelper
    {
        [Test]
        public static void SaveAsMD()
        {
            //ExStart:SaveAsMD
            DocumentBuilder builder = new DocumentBuilder();
            builder.Writeln("Some text!");

            // specify MarkDownSaveOptions
            Saving.MarkdownSaveOptions saveOptions = (Saving.MarkdownSaveOptions)SaveOptions.CreateSaveOptions(SaveFormat.Markdown);
            
            builder.Document.Save(ArtifactsDir + "TestDocument.md", saveOptions);
            //ExEnd:SaveAsMD
        }

        [Test]
        public static void ExportIntoMarkdownWithTableContentAlignment()
        {
            // ExStart:ExportIntoMarkdownWithTableContentAlignment
            DocumentBuilder builder = new DocumentBuilder();

            // Create a new table with two cells.
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
            builder.Write("Cell1");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.Write("Cell2");

            Saving.MarkdownSaveOptions saveOptions = new Saving.MarkdownSaveOptions();
            // Makes all paragraphs inside table to be aligned to Left. 
            saveOptions.TableContentAlignment = TableContentAlignment.Left;
            builder.Document.Save(ArtifactsDir + "left.md", saveOptions);

            // Makes all paragraphs inside table to be aligned to Right. 
            saveOptions.TableContentAlignment = TableContentAlignment.Right;
            builder.Document.Save(ArtifactsDir + "right.md", saveOptions);

            // Makes all paragraphs inside table to be aligned to Center. 
            saveOptions.TableContentAlignment = TableContentAlignment.Center;
            builder.Document.Save(ArtifactsDir + "center.md", saveOptions);

            // Makes all paragraphs inside table to be aligned automatically.
            // The alignment in this case will be taken from the first paragraph in corresponding table column.
            saveOptions.TableContentAlignment = TableContentAlignment.Auto;
            builder.Document.Save(ArtifactsDir + "auto.md", saveOptions);
            // ExEnd:ExportIntoMarkdownWithTableContentAlignment
        }
    }
}
