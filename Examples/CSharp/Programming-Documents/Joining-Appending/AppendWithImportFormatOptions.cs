using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class AppendWithImportFormatOptions : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:AppendWithImportFormatOptions
            Document srcDoc = new Document(JoiningAppendingDir + "TestFile.SourceList.doc");
            Document dstDoc = new Document(JoiningAppendingDir + "TestFile.DestinationList.doc");

            ImportFormatOptions options = new ImportFormatOptions();
            // Specify that if numbering clashes in source and destination documents,
            // then a numbering from the source document will be used
            options.KeepSourceNumbering = true;

            dstDoc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles, options);
            //ExEnd:AppendWithImportFormatOptions
        }
    }
}