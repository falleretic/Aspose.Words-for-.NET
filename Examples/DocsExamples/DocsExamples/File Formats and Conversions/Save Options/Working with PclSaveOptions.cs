﻿using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace DocsExamples.File_Formats_and_Conversions.Save_Options
{
    internal class WorkingWithPclSaveOptions : DocsExamplesBase
    {
        [Test]
        public static void RasterizeTransformedElements()
        {
            //ExStart:RasterizeTransformedElements
            Document doc = new Document(MyDir + "Rendering.docx");

            PclSaveOptions saveOptions = new PclSaveOptions
            {
                SaveFormat = SaveFormat.Pcl, RasterizeTransformedElements = false
            };

            doc.Save(ArtifactsDir + "WorkingWithPclSaveOptions.RasterizeTransformedElements.pcl", saveOptions);
            //ExEnd:RasterizeTransformedElements
        }
    }
}