﻿using Aspose.Words.Saving;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class WorkingWithRtfSaveOptions : TestDataHelper
    {
        [Test]
        public static void SavingImagesAsWmf()
        {
            //ExStart:SavingImagesAsWmf
            Document doc = new Document(DocumentDir + "TestFile.doc");

            RtfSaveOptions saveOpts = new RtfSaveOptions();
            saveOpts.SaveImagesAsWmf = true;

            doc.Save(ArtifactsDir + "output.rtf", saveOpts);
            //ExEnd:SavingImagesAsWmf
        }
    }
}