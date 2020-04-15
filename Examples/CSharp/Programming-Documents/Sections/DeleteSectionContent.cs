﻿using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Sections
{
    class DeleteSectionContent : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:DeleteSectionContent
            Document doc = new Document(SectionsDir + "Document.doc");
            
            Section section = doc.Sections[0];
            section.ClearContent();
            //ExEnd:DeleteSectionContent
        }
    }
}