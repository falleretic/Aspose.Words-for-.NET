using Aspose.Words.Replacing;
using System;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Find_and_Replace
{
    class FindReplaceUsingMetaCharacters : TestDataHelper
    {
        public static void Run()
        {
            /* meta-characters
            &p - paragraph break
            &b - section break
            &m - page break
            &l - manual line break
            */

            MetaCharactersInSearchPattern();
            ReplaceTextContainingMetaCharacters();
        }

        public static void MetaCharactersInSearchPattern()
        {
            //ExStart:MetaCharactersInSearchPattern
            Document doc = new Document();

            // Use a document builder to add content to the document
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("This is Line 1");
            builder.Writeln("This is Line 2");

            var findReplaceOptions = new FindReplaceOptions();

            doc.Range.Replace("This is Line 1&pThis is Line 2", "This is replaced line", findReplaceOptions);

            builder.MoveToDocumentEnd();
            builder.Write("This is Line 1");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("This is Line 2");

            doc.Range.Replace("This is Line 1&mThis is Line 2", "Page break is replaced with new text.",
                findReplaceOptions);

            doc.Save(ArtifactsDir + "MetaCharactersInSearchPattern.docx");
            //ExEnd:MetaCharactersInSearchPattern
        }

        public static void ReplaceTextContainingMetaCharacters()
        {
            //ExStart:ReplaceTextContaingMetaCharacters
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Font.Name = "Arial";
            builder.Writeln("First section");
            builder.Writeln("  1st paragraph");
            builder.Writeln("  2nd paragraph");
            builder.Writeln("{insert-section}");
            builder.Writeln("Second section");
            builder.Writeln("  1st paragraph");

            FindReplaceOptions options = new FindReplaceOptions();
            options.ApplyParagraphFormat.Alignment = ParagraphAlignment.Center;

            // Double each paragraph break after word "section", add kind of underline and make it centered.
            int count = doc.Range.Replace("section&p", "section&p----------------------&p", options);

            // Insert section break instead of custom text tag.
            count = doc.Range.Replace("{insert-section}", "&b", options);

            doc.Save(ArtifactsDir + "ReplaceTextContainingMetaCharacters.docx");
            //ExEnd:ReplaceTextContaingMetaCharacters
        }
    }
}