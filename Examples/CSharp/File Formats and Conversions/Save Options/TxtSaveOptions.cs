using Aspose.Words.Saving;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class WorkingWithTxt : TestDataHelper
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
            
            TxtSaveOptions saveOptions = new TxtSaveOptions();
            saveOptions.AddBidiMarks = true;

            doc.Save(ArtifactsDir + "AddBidiMarks.txt", saveOptions);
            //ExEnd:AddBidiMarks
        }

        [Test]
        public static void ExportHeadersFootersMode()
        {
            //ExStart:ExportHeadersFootersMode
            Document doc = new Document();

            // Insert even and primary headers/footers into the document
            // The primary header/footers should override the even ones 
            doc.FirstSection.HeadersFooters.Add(new HeaderFooter(doc, HeaderFooterType.HeaderEven));
            doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderEven].AppendParagraph("Even header");
            doc.FirstSection.HeadersFooters.Add(new HeaderFooter(doc, HeaderFooterType.FooterEven));
            doc.FirstSection.HeadersFooters[HeaderFooterType.FooterEven].AppendParagraph("Even footer");
            doc.FirstSection.HeadersFooters.Add(new HeaderFooter(doc, HeaderFooterType.HeaderPrimary));
            doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary].AppendParagraph("Primary header");
            doc.FirstSection.HeadersFooters.Add(new HeaderFooter(doc, HeaderFooterType.FooterPrimary));
            doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary].AppendParagraph("Primary footer");

            // Insert pages that would display these headers and footers
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Page 1");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 2");
            builder.InsertBreak(BreakType.PageBreak); 
            builder.Write("Page 3");

            TxtSaveOptions options = new TxtSaveOptions();
            options.SaveFormat = SaveFormat.Text;
            // All headers and footers are placed at the very end of the output document
            options.ExportHeadersFootersMode = TxtExportHeadersFootersMode.AllAtEnd;
            
            doc.Save(ArtifactsDir + "ExportHeadersFootersModeA.txt", options);

            // Only primary headers and footers are exported at the beginning and end of each section
            options.ExportHeadersFootersMode = TxtExportHeadersFootersMode.PrimaryOnly;
            
            doc.Save(ArtifactsDir + "ExportHeadersFootersModeB.txt", options);

            // No headers and footers are exported
            options.ExportHeadersFootersMode = TxtExportHeadersFootersMode.None;
            
            doc.Save(ArtifactsDir + "ExportHeadersFootersModeC.txt", options);
            //ExEnd:ExportHeadersFootersMode
        }

        [Test]
        public static void UseTabCharacterPerLevelForListIndentation()
        {
            //ExStart:UseTabCharacterPerLevelForListIndentation
            Document doc = new Document(LoadingSavingDir + "List indentation.docx");

            TxtSaveOptions options = new TxtSaveOptions();
            options.ListIndentation.Count = 1;
            options.ListIndentation.Character = '\t';

            doc.Save(ArtifactsDir + "UseTabCharacterPerLevelForListIndentation.txt", options);
            //ExEnd:UseTabCharacterPerLevelForListIndentation
        }

        [Test]
        public static void UseSpaceCharacterPerLevelForListIndentation()
        {
            //ExStart:UseSpaceCharacterPerLevelForListIndentation
            Document doc = new Document(LoadingSavingDir + "List indentation.docx");

            TxtSaveOptions options = new TxtSaveOptions();
            options.ListIndentation.Count = 3;
            options.ListIndentation.Character = ' ';

            doc.Save(ArtifactsDir + "UseSpaceCharacterPerLevelForListIndentation.txt", options);
            //ExEnd:UseSpaceCharacterPerLevelForListIndentation
        }

        [Test]
        public static void DefaultLevelForListIndentation()
        {
            //ExStart:DefaultLevelForListIndentation
            Document doc1 = new Document(LoadingSavingDir + "List indentation.docx");
            doc1.Save(ArtifactsDir + "DefaultLevelForListIndentation1.txt");

            Document doc2 = new Document(LoadingSavingDir + "List indentation.docx");
            TxtSaveOptions options = new TxtSaveOptions();
            doc2.Save(ArtifactsDir + "DefaultLevelForListIndentation2.txt", options);
            //ExEnd:DefaultLevelForListIndentation
        }
    }
}