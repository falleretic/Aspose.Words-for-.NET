﻿using Aspose.Words.Saving;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.File_Formats_and_Conversions.Save_Options
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

            doc.Save(ArtifactsDir + "TxtSaveOptions.AddBidiMarks.txt", saveOptions);
            //ExEnd:AddBidiMarks
        }

        [Test]
        public static void ExportHeadersFootersMode()
        {
            //ExStart:ExportHeadersFootersMode
            Document doc = new Document();

            // Insert even and primary headers/footers into the document.
            // The primary header/footers should override the even ones.
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

            TxtSaveOptions saveOptions = new TxtSaveOptions();
            saveOptions.SaveFormat = SaveFormat.Text;
            // All headers and footers are placed at the very end of the output document.
            saveOptions.ExportHeadersFootersMode = TxtExportHeadersFootersMode.AllAtEnd;
            
            doc.Save(ArtifactsDir + "TxtSaveOptions.ExportHeadersFootersModeA.txt", saveOptions);

            // Only primary headers and footers are exported at the beginning and end of each section.
            saveOptions.ExportHeadersFootersMode = TxtExportHeadersFootersMode.PrimaryOnly;
            
            doc.Save(ArtifactsDir + "TxtSaveOptions.ExportHeadersFootersModeB.txt", saveOptions);

            // No headers and footers are exported.
            saveOptions.ExportHeadersFootersMode = TxtExportHeadersFootersMode.None;
            
            doc.Save(ArtifactsDir + "TxtSaveOptions.ExportHeadersFootersModeC.txt", saveOptions);
            //ExEnd:ExportHeadersFootersMode
        }

        [Test]
        public static void UseTabCharacterPerLevelForListIndentation()
        {
            //ExStart:UseTabCharacterPerLevelForListIndentation
            Document doc = new Document(MyDir + "List indentation.docx");

            TxtSaveOptions saveOptions = new TxtSaveOptions();
            saveOptions.ListIndentation.Count = 1;
            saveOptions.ListIndentation.Character = '\t';

            doc.Save(ArtifactsDir + "TxtSaveOptions.UseTabCharacterPerLevelForListIndentation.txt", saveOptions);
            //ExEnd:UseTabCharacterPerLevelForListIndentation
        }

        [Test]
        public static void UseSpaceCharacterPerLevelForListIndentation()
        {
            //ExStart:UseSpaceCharacterPerLevelForListIndentation
            Document doc = new Document(MyDir + "List indentation.docx");

            TxtSaveOptions saveOptions = new TxtSaveOptions();
            saveOptions.ListIndentation.Count = 3;
            saveOptions.ListIndentation.Character = ' ';

            doc.Save(ArtifactsDir + "TxtSaveOptions.UseSpaceCharacterPerLevelForListIndentation.txt", saveOptions);
            //ExEnd:UseSpaceCharacterPerLevelForListIndentation
        }
    }
}