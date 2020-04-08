namespace Aspose.Words.Examples.CSharp.Programming_Documents.Joining_Appending
{
    class AppendWithImportFormatOptions : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:AppendWithImportFormatOptions
            Document srcDoc = new Document(JoiningAppendingDir + "source.docx");
            Document dstDoc = new Document(JoiningAppendingDir + "destination.docx");

            ImportFormatOptions options = new ImportFormatOptions();
            // Specify that if numbering clashes in source and destination documents,
            // then a numbering from the source document will be used
            options.KeepSourceNumbering = true;

            dstDoc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles, options);
            //ExEnd:AppendWithImportFormatOptions
        }
    }
}