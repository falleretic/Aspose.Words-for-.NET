using System;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class CheckDMLTextEffect : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:CheckDMLTextEffect
            Document doc = new Document(DocumentDir + "Document.doc");
            
            RunCollection runs = doc.FirstSection.Body.FirstParagraph.Runs;
            Font runFont = runs[0].Font;

            // One run might have several Dml text effects applied
            Console.WriteLine(runFont.HasDmlEffect(TextDmlEffect.Shadow));
            Console.WriteLine(runFont.HasDmlEffect(TextDmlEffect.Effect3D));
            Console.WriteLine(runFont.HasDmlEffect(TextDmlEffect.Reflection));
            Console.WriteLine(runFont.HasDmlEffect(TextDmlEffect.Outline));
            Console.WriteLine(runFont.HasDmlEffect(TextDmlEffect.Fill));
            //ExEnd:CheckDMLTextEffect
        }
    }
}