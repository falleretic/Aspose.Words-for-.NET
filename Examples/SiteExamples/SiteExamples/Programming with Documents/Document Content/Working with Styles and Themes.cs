using System;
using System.Drawing;
using Aspose.Words;
using NUnit.Framework;

namespace SiteExamples.Programming_with_Documents.Document_Content
{
    class WorkingWithStylesAndThemes : SiteExamplesBase
    {
        [Test]
        public static void AccessStyles()
        {
            //ExStart:AccessStyles
            Document doc = new Document();

            // Get styles collection from document
            StyleCollection styles = doc.Styles;
            string styleName = "";

            // Iterate through all the styles
            foreach (Style style in styles)
            {
                if (styleName == "")
                {
                    styleName = style.Name;
                    Console.WriteLine(styleName);
                }
                else
                {
                    styleName = styleName + ", " + style.Name;
                    Console.WriteLine(styleName);
                }
            }
            //ExEnd:AccessStyles
        }

        [Test]
        public static void CopyStylesFromDocument()
        {
            //ExStart:CopyStylesFromDocument
            Document doc = new Document();
            Document target = new Document(MyDir + "Rendering.docx");

            target.CopyStylesFromTemplate(doc);

            doc.Save(ArtifactsDir + "CopyStyles.docx");
            //ExEnd:CopyStylesFromDocument
        }

        /// <summary>
        ///  Shows how to get theme properties.
        /// </summary>
        [Test]
        public static void GetThemeProperties()
        {
            //ExStart:GetThemeProperties
            Document doc = new Document();

            Aspose.Words.Themes.Theme theme = doc.Theme;
            // Major (Headings) font for Latin characters
            Console.WriteLine(theme.MajorFonts.Latin);
            // Minor (Body) font for EastAsian characters
            Console.WriteLine(theme.MinorFonts.EastAsian);
            // Color for theme color Accent 1
            Console.WriteLine(theme.Colors.Accent1);
            //ExEnd:GetThemeProperties 
        }

        /// <summary>
        ///  Shows how to set theme properties.
        /// </summary>
        [Test]
        public static void SetThemeProperties()
        {
            // ExStart:SetThemeProperties
            Document doc = new Document();

            Aspose.Words.Themes.Theme theme = doc.Theme;
            // Set Times New Roman font as Body theme font for Latin Character
            theme.MinorFonts.Latin = "Times New Roman";
            // Set Color.Gold for theme color Hyperlink
            theme.Colors.Hyperlink = Color.Gold;
            // ExEnd:SetThemeProperties 
        }

        [Test]
        public static void ParagraphInsertStyleSeparator()
        {
            //ExStart:ParagraphInsertStyleSeparator
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Style paraStyle = builder.Document.Styles.Add(StyleType.Paragraph, "MyParaStyle");
            paraStyle.Font.Bold = false;
            paraStyle.Font.Size = 8;
            paraStyle.Font.Name = "Arial";

            // Append text with "Heading 1" style
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Write("Heading 1");
            builder.InsertStyleSeparator();

            // Append text with another style
            builder.ParagraphFormat.StyleName = paraStyle.Name;
            builder.Write("This is text with some other formatting ");

            doc.Save(ArtifactsDir + "InsertStyleSeparator.docx");
            //ExEnd:ParagraphInsertStyleSeparator
        }
    }
}