using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.DocumentEx
{
    class WorkingWithImportFormatOptions : TestDataHelper
    {
        [Test]
        public static void SmartStyleBehavior()
        {
            //ExStart:SmartStyleBehavior
            Document srcDoc = new Document(JoiningAppendingDir + "TestFile.Source.doc");
            Document dstDoc = new Document(JoiningAppendingDir + "TestFile.Destination.doc");

            DocumentBuilder builder = new DocumentBuilder(dstDoc);
            builder.MoveToDocumentEnd();
            builder.InsertBreak(BreakType.PageBreak);

            ImportFormatOptions options = new ImportFormatOptions();
            options.SmartStyleBehavior = true;
            builder.InsertDocument(srcDoc, ImportFormatMode.UseDestinationStyles, options);
            //ExEnd:SmartStyleBehavior
        }

        [Test]
        public static void KeepSourceNumbering()
        {
            //ExStart:KeepSourceNumbering
            Document srcDoc = new Document(JoiningAppendingDir + "TestFile.Source.doc");
            Document dstDoc = new Document(JoiningAppendingDir + "TestFile.Destination.doc");

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

        [Test]
        public static void IgnoreTextBoxes()
        {
            //ExStart:IgnoreTextBoxes
            Document srcDoc = new Document(JoiningAppendingDir + "TestFile.Source.doc");
            Document dstDoc = new Document(JoiningAppendingDir + "TestFile.Destination.doc");

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

        [Test]
        public static void IgnoreHeaderFooter()
        {
            // ExStart:IgnoreHeaderFooter
            Document srcDocument = new Document(JoiningAppendingDir + "source.docx");
            Document dstDocument = new Document(JoiningAppendingDir + "destination.docx");

            ImportFormatOptions importFormatOptions = new ImportFormatOptions();
            importFormatOptions.IgnoreHeaderFooter = false;

            dstDocument.AppendDocument(srcDocument, ImportFormatMode.KeepSourceFormatting, importFormatOptions);
            dstDocument.Save(ArtifactsDir + "IgnoreHeaderFooter.docx");
            // ExEnd:IgnoreHeaderFooter
        }
    }
}