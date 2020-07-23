using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class AppendWithImportFormatOptions : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:AppendWithImportFormatOptions
            Document srcDoc = new Document(JoiningAppendingDir + "Document source with list.docx");
            Document dstDoc = new Document(JoiningAppendingDir + "Document destination with list.docx");

            ImportFormatOptions options = new ImportFormatOptions();
            // Specify that if numbering clashes in source and destination documents,
            // then a numbering from the source document will be used
            options.KeepSourceNumbering = true;

            dstDoc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles, options);
            //ExEnd:AppendWithImportFormatOptions
        }
    }
}