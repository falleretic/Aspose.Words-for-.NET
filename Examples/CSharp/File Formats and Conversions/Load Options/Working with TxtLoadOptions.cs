using System;
using System.IO;
using System.Text;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.File_Formats_and_Conversions.Load_Options
{
    internal class WorkingWithTxtLoadOptions : TestDataHelper
    {
        [Test, Description("Shows how to convert numbered list items from plain text format.")]
        public void DetectNumberingWithWhitespaces()
        {
            //ExStart:DetectNumberingWithWhitespaces
            TxtLoadOptions loadOptions = new TxtLoadOptions();
            loadOptions.DetectNumberingWithWhitespaces = false;

            Document doc = new Document(MyDir + "Txt document.txt", loadOptions);
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
