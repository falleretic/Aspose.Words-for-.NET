using System;
using System.IO;
using System.Text;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class TxtLoadOptions : TestDataHelper
    {
        [Test]
        public static void DetectNumberingWithWhitespaces()
        {
            //ExStart:DetectNumberingWithWhitespaces
            Words.TxtLoadOptions loadOptions = new Words.TxtLoadOptions();
            loadOptions.DetectNumberingWithWhitespaces = false;

            Document doc = new Document(LoadingSavingDir + "Txt document.txt", loadOptions);
            doc.Save(ArtifactsDir + "DetectNumberingWithWhitespaces.docx");
            //ExEnd:DetectNumberingWithWhitespaces
        }

        [Test]
        public static void HandleSpacesOptions()
        {
            //ExStart:HandleSpacesOptions
            string textDoc = "      Line 1 \n" +
                             "    Line 2   \n" +
                             " Line 3       ";
            
            Words.TxtLoadOptions loadOptions = new Words.TxtLoadOptions();
            loadOptions.LeadingSpacesOptions = TxtLeadingSpacesOptions.Trim;
            loadOptions.TrailingSpacesOptions = TxtTrailingSpacesOptions.Trim;
            
            Document doc = new Document(new MemoryStream(Encoding.UTF8.GetBytes(textDoc)), loadOptions);
            doc.Save(ArtifactsDir + "HandleSpacesOptions.docx");
            //ExEnd:HandleSpacesOptions
        }

        [Test]
        public static void DocumentTextDirection()
        {
            //ExStart:DocumentTextDirection
            Words.TxtLoadOptions loadOptions = new Words.TxtLoadOptions();
            loadOptions.DocumentDirection = DocumentDirection.Auto;

            Document doc = new Document(LoadingSavingDir + "Hebrew text.txt", loadOptions);

            Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;
            Console.WriteLine(paragraph.ParagraphFormat.Bidi);

            doc.Save(ArtifactsDir + "DocumentTextDirection.docx");
            //ExEnd:DocumentTextDirection
        }
    }
}
