﻿using System;
using Aspose.Words.Themes;
using System.Drawing;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Theme
{
    class ManipulateThemeProperties : TestDataHelper
    {
        public static void Run()
        {
            GetThemeProperties();
            SetThemeProperties();
        }

        /// <summary>
        ///  Shows how to get theme properties.
        /// </summary>             
        private static void GetThemeProperties()
        {
            //ExStart:GetThemeProperties
            Document doc = new Document(ThemeDir + "Document.doc");
            
            Theme theme = doc.Theme;
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
        private static void SetThemeProperties()
        {
            // ExStart:SetThemeProperties
            Document doc = new Document(ThemeDir + "Document.doc");
            
            Theme theme = doc.Theme;
            // Set Times New Roman font as Body theme font for Latin Character
            theme.MinorFonts.Latin = "Times New Roman";
            // Set Color.Gold for theme color Hyperlink
            theme.Colors.Hyperlink = Color.Gold;
            // ExEnd:SetThemeProperties 
        }
    }
}