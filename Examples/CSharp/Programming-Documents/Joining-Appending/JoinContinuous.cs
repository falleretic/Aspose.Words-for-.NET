﻿namespace Aspose.Words.Examples.CSharp.Programming_Documents.Joining_and_Appending
{
    internal class JoinContinuous : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:JoinContinuous
            Document dstDoc = new Document(JoiningAppendingDir + "TestFile.Destination.doc");
            Document srcDoc = new Document(JoiningAppendingDir + "TestFile.Source.doc");

            // Make the document appear straight after the destination documents content
            srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.Continuous;

            // Append the source document using the original styles found in the source document
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            
            dstDoc.Save(ArtifactsDir + "JoinContinuous.docx");
            //ExEnd:JoinContinuous
        }
    }
}