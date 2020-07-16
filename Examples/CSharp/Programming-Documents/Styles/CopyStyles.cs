﻿using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Styles
{
    class CopyStyles : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:CopyStylesFromDocument
            Document doc = new Document(StyleDir + "template.docx");
            Document target = new Document(StyleDir + "TestFile.doc");
            
            target.CopyStylesFromTemplate(doc);
            
            doc.Save(ArtifactsDir + "CopyStyles.docx");
            //ExEnd:CopyStylesFromDocument
        }
    }
}