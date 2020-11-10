using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace DocsExamples.File_Formats_and_Conversions.Save_Options
{
    internal class WorkingWithTxtSaveOptions : DocsExamplesBase
    {
        [Test]
        public static void AddBidiMarks()
        {
            //ExStart:AddBidiMarks
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello world!");
            builder.ParagraphFormat.Bidi = true;
            builder.Writeln("שלום עולם!");
            builder.Writeln("مرحبا بالعالم!");

            TxtSaveOptions txtSaveOptions = new TxtSaveOptions { AddBidiMarks = true };

            doc.Save(ArtifactsDir + "WorkingWithTxtSaveOptions.AddBidiMarks.txt", txtSaveOptions);
            //ExEnd:AddBidiMarks
        }

        [Test]
        public static void ExportHeadersFootersMode()
        {
            //ExStart:ExportHeadersFootersMode
            Document doc = new Document();

            // Insert even and primary headers/footers into the document.
            // The primary header/footers should override the even ones,
            doc.FirstSection.HeadersFooters.Add(new HeaderFooter(doc, HeaderFooterType.HeaderEven));
            doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderEven].AppendParagraph("Even header");
            doc.FirstSection.HeadersFooters.Add(new HeaderFooter(doc, HeaderFooterType.FooterEven));
            doc.FirstSection.HeadersFooters[HeaderFooterType.FooterEven].AppendParagraph("Even footer");
            doc.FirstSection.HeadersFooters.Add(new HeaderFooter(doc, HeaderFooterType.HeaderPrimary));
            doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary].AppendParagraph("Primary header");
            doc.FirstSection.HeadersFooters.Add(new HeaderFooter(doc, HeaderFooterType.FooterPrimary));
            doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary].AppendParagraph("Primary footer");

            // Insert pages that would display these headers and footers.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Page 1");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 2");
            builder.InsertBreak(BreakType.PageBreak); 
            builder.Write("Page 3");

            // All headers and footers are placed at the very end of the output document.
            TxtSaveOptions txtSaveOptions = new TxtSaveOptions
            {
                SaveFormat = SaveFormat.Text, ExportHeadersFootersMode = TxtExportHeadersFootersMode.AllAtEnd
            };
            doc.Save(ArtifactsDir + "WorkingWithTxtSaveOptions.ExportHeadersFootersAllAtEnd.txt", txtSaveOptions);

            // Only primary headers and footers are exported at the beginning and end of each section.
            txtSaveOptions.ExportHeadersFootersMode = TxtExportHeadersFootersMode.PrimaryOnly;
            doc.Save(ArtifactsDir + "WorkingWithTxtSaveOptions.ExportHeadersFootersPrimaryOnly.txt", txtSaveOptions);

            // No headers and footers are exported.
            txtSaveOptions.ExportHeadersFootersMode = TxtExportHeadersFootersMode.None;
            doc.Save(ArtifactsDir + "WorkingWithTxtSaveOptions.DoNotExportHeadersFooters.txt", txtSaveOptions);
            //ExEnd:ExportHeadersFootersMode
        }

        [Test]
        public static void UseTabCharacterPerLevelForListIndentation()
        {
            //ExStart:UseTabCharacterPerLevelForListIndentation
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a list with three levels of indentation.
            builder.ListFormat.ApplyNumberDefault();
            builder.Writeln("Item 1");
            builder.ListFormat.ListIndent();
            builder.Writeln("Item 2");
            builder.ListFormat.ListIndent(); 
            builder.Write("Item 3");

            TxtSaveOptions txtSaveOptions = new TxtSaveOptions();
            txtSaveOptions.ListIndentation.Count = 1;
            txtSaveOptions.ListIndentation.Character = '\t';

            doc.Save(ArtifactsDir + "WorkingWithTxtSaveOptions.UseTabCharacterPerLevelForListIndentation.txt", txtSaveOptions);
            //ExEnd:UseTabCharacterPerLevelForListIndentation
        }

        [Test]
        public static void UseSpaceCharacterPerLevelForListIndentation()
        {
            //ExStart:UseSpaceCharacterPerLevelForListIndentation
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a list with three levels of indentation.
            builder.ListFormat.ApplyNumberDefault();
            builder.Writeln("Item 1");
            builder.ListFormat.ListIndent();
            builder.Writeln("Item 2");
            builder.ListFormat.ListIndent(); 
            builder.Write("Item 3");

            TxtSaveOptions txtSaveOptions = new TxtSaveOptions();
            txtSaveOptions.ListIndentation.Count = 3;
            txtSaveOptions.ListIndentation.Character = ' ';

            doc.Save(ArtifactsDir + "WorkingWithTxtSaveOptions.UseSpaceCharacterPerLevelForListIndentation.txt", txtSaveOptions);
            //ExEnd:UseSpaceCharacterPerLevelForListIndentation
        }
    }
}