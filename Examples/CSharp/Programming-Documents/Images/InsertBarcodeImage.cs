namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Images
{
    class InsertBarcodeImage : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:InsertBarcodeImage
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // The number of pages the document should have
            const int numPages = 4;
            // The document starts with one section, insert the barcode into this existing section
            InsertBarcodeIntoFooter(builder, doc.FirstSection, HeaderFooterType.FooterPrimary);

            for (int i = 1; i < numPages; i++)
            {
                // Clone the first section and add it into the end of the document
                Section cloneSection = (Section) doc.FirstSection.Clone(false);
                cloneSection.PageSetup.SectionStart = SectionStart.NewPage;
                doc.AppendChild(cloneSection);

                // Insert the barcode and other information into the footer of the section
                InsertBarcodeIntoFooter(builder, cloneSection, HeaderFooterType.FooterPrimary);
            }

            // Save the document as a PDF to disk
            // You can also save this directly to a stream
            doc.Save(ArtifactsDir + "InsertBarcodeImage.docx");
            //ExEnd:InsertBarcodeImage
        }

        //ExStart:InsertBarcodeIntoFooter
        private static void InsertBarcodeIntoFooter(DocumentBuilder builder, Section section,
            HeaderFooterType footerType)
        {
            // Move to the footer type in the specific section
            builder.MoveToSection(section.Document.IndexOf(section));
            builder.MoveToHeaderFooter(footerType);

            // Insert the barcode, then move to the next line and insert the ID along with the page number
            // Use pageId if you need to insert a different barcode on each page. 0 = First page, 1 = Second page etc.
            builder.InsertImage(
                System.Drawing.Image.FromFile(RunExamples.GetDataDir_WorkingWithImages() + "Barcode1.png"));
            builder.Writeln();
            builder.Write("1234567890");
            builder.InsertField("PAGE");

            // Create a right aligned tab at the right margin
            double tabPos = section.PageSetup.PageWidth - section.PageSetup.RightMargin - section.PageSetup.LeftMargin;
            builder.CurrentParagraph.ParagraphFormat.TabStops.Add(new TabStop(tabPos, TabAlignment.Right,
                TabLeader.None));

            // Move to the right hand side of the page and insert the page and page total
            builder.Write(ControlChar.Tab);
            builder.InsertField("PAGE");
            builder.Write(" of ");
            builder.InsertField("NUMPAGES");
        }
        //ExEnd:InsertBarcodeIntoFooter
    }
}