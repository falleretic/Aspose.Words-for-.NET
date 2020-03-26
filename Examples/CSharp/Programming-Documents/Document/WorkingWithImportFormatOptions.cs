namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class WorkingWithImportFormatOptions : TestDataHelper
    {
        public static void Run()
        {
            SmartStyleBehavior();
            KeepSourceNumbering();
            IgnoreTextBoxes();
        }

        static void SmartStyleBehavior()
        {
            //ExStart:SmartStyleBehavior
            Document srcDoc = new Document(DocumentDir + "source.docx");
            Document dstDoc = new Document(DocumentDir + "destination.docx");

            DocumentBuilder builder = new DocumentBuilder(dstDoc);
            builder.MoveToDocumentEnd();
            builder.InsertBreak(BreakType.PageBreak);

            ImportFormatOptions options = new ImportFormatOptions();
            options.SmartStyleBehavior = true;
            builder.InsertDocument(srcDoc, ImportFormatMode.UseDestinationStyles, options);
            //ExEnd:SmartStyleBehavior
        }

        static void KeepSourceNumbering()
        {
            //ExStart:KeepSourceNumbering
            Document srcDoc = new Document(DocumentDir + "source.docx");
            Document dstDoc = new Document(DocumentDir + "destination.docx");

            ImportFormatOptions importFormatOptions = new ImportFormatOptions();
            // Keep source list formatting when importing numbered paragraphs
            importFormatOptions.KeepSourceNumbering = true;
            
            NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting,
                importFormatOptions);

            ParagraphCollection srcParas = srcDoc.FirstSection.Body.Paragraphs;
            foreach (Paragraph srcPara in srcParas)
            {
                Node importedNode = importer.ImportNode(srcPara, false);
                dstDoc.FirstSection.Body.AppendChild(importedNode);
            }

            dstDoc.Save(ArtifactsDir + "output.docx");
            //ExEnd:KeepSourceNumbering
        }

        public static void IgnoreTextBoxes()
        {
            //ExStart:IgnoreTextBoxes
            Document srcDoc = new Document(DocumentDir + "source.docx");
            Document dstDoc = new Document(DocumentDir + "destination.docx");

            ImportFormatOptions importFormatOptions = new ImportFormatOptions();
            // Keep the source text boxes formatting when importing
            importFormatOptions.IgnoreTextBoxes = false;
            
            NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting,
                importFormatOptions);

            ParagraphCollection srcParas = srcDoc.FirstSection.Body.Paragraphs;
            foreach (Paragraph srcPara in srcParas)
            {
                Node importedNode = importer.ImportNode(srcPara, true);
                dstDoc.FirstSection.Body.AppendChild(importedNode);
            }

            dstDoc.Save(ArtifactsDir + "output.docx");
            //ExEnd:IgnoreTextBoxes
        }
    }
}