using System;
using System.IO;
using System.Text;
using Aspose.Words;
using NUnit.Framework;

namespace SiteExamples.File_Formats_and_Conversions.Load_Options
{
    internal class WorkingWithTxtLoadOptions : SiteExamplesBase
    {
        [Test, Description("Shows how to convert numbered list items from plain text format.")]
        public void DetectNumberingWithWhitespaces()
        {
            //ExStart:DetectNumberingWithWhitespaces
            // Create a plaintext document in the form of a string with parts that may be interpreted as lists.
            // Upon loading, the first three lists will always be detected by Aspose.Words,
            // and List objects will be created for them after loading.
            const string textDoc = "Full stop delimiters:\n" +
                                   "1. First list item 1\n" +
                                   "2. First list item 2\n" +
                                   "3. First list item 3\n\n" +
                                   "Right bracket delimiters:\n" +
                                   "1) Second list item 1\n" +
                                   "2) Second list item 2\n" +
                                   "3) Second list item 3\n\n" +
                                   "Bullet delimiters:\n" +
                                   "• Third list item 1\n" +
                                   "• Third list item 2\n" +
                                   "• Third list item 3\n\n" +
                                   "Whitespace delimiters:\n" +
                                   "1 Fourth list item 1\n" +
                                   "2 Fourth list item 2\n" +
                                   "3 Fourth list item 3";

            // The fourth list, with whitespace inbetween the list number and list item contents,
            // will only be detected as a list if "DetectNumberingWithWhitespaces" in a LoadOptions object is set to true,
            // to avoid paragraphs that start with numbers being mistakenly detected as lists.
            TxtLoadOptions loadOptions = new TxtLoadOptions();
            loadOptions.DetectNumberingWithWhitespaces = true;

            // Load the document while applying LoadOptions as a parameter and verify the result.
            Document doc = new Document(new MemoryStream(Encoding.UTF8.GetBytes(textDoc)), loadOptions);
            doc.Save(ArtifactsDir + "TxtLoadOptions.DetectNumberingWithWhitespaces.docx");
            //ExEnd:DetectNumberingWithWhitespaces
        }

        [Test, Description("Shows how to handle leading and trailing spaces.")]
        public void HandleSpacesOptions()
        {
            //ExStart:HandleSpacesOptions
            const string textDoc = "      Line 1 \n" +
                                   "    Line 2   \n" +
                                   " Line 3       ";
            
            TxtLoadOptions loadOptions = new TxtLoadOptions();
            loadOptions.LeadingSpacesOptions = TxtLeadingSpacesOptions.Trim;
            loadOptions.TrailingSpacesOptions = TxtTrailingSpacesOptions.Trim;
            
            Document doc = new Document(new MemoryStream(Encoding.UTF8.GetBytes(textDoc)), loadOptions);
            doc.Save(ArtifactsDir + "TxtLoadOptions.HandleSpacesOptions.docx");
            //ExEnd:HandleSpacesOptions
        }

        [Test, Description("Shows how to specify text direction.")]
        public void DocumentTextDirection()
        {
            //ExStart:DocumentTextDirection
            TxtLoadOptions loadOptions = new TxtLoadOptions();
            loadOptions.DocumentDirection = DocumentDirection.Auto;

            Document doc = new Document(MyDir + "Hebrew text.txt", loadOptions);

            Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;
            Console.WriteLine(paragraph.ParagraphFormat.Bidi);

            doc.Save(ArtifactsDir + "TxtLoadOptions.DocumentTextDirection.docx");
            //ExEnd:DocumentTextDirection
        }
    }
}
