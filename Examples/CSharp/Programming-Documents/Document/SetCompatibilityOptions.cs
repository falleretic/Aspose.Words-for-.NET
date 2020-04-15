﻿using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class SetCompatibilityOptions : TestDataHelper
    {
        [Test]
        public static void OptimizeFor()
        {
            //ExStart:OptimizeFor
            Document doc = new Document(DocumentDir + "TestFile.docx");
            doc.CompatibilityOptions.OptimizeFor(Settings.MsWordVersion.Word2016);

            doc.Save(ArtifactsDir + "TestFile.docx");
            //ExEnd:OptimizeFor
        }
    }
}