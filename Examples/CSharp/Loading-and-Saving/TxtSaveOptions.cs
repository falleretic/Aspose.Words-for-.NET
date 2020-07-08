using System;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class WorkingWithTxt : TestDataHelper
    {
        [Test]
        public static void SaveAsTxt()
        {
            //ExStart:SaveAsTxt
            Document doc = new Document(LoadingSavingDir + "Document.doc");
            doc.Save(ArtifactsDir + "SaveAsTxt.txt");
            //ExEnd:SaveAsTxt
        }

        [Test]
        public static void AddBidiMarks()
        {
            //ExStart:AddBidiMarks
            Document doc = new Document(LoadingSavingDir + "Input.docx");
            
            TxtSaveOptions saveOptions = new TxtSaveOptions();
            saveOptions.AddBidiMarks = true;

            doc.Save(ArtifactsDir + "AddBidiMarks.txt", saveOptions);
            //ExEnd:AddBidiMarks
        }

        [Test]
        public static void ExportHeadersFootersMode()
        {
            //ExStart:ExportHeadersFootersMode
            Document doc = new Document(LoadingSavingDir + "TxtExportHeadersFootersMode.docx");

            TxtSaveOptions options = new TxtSaveOptions();
            options.SaveFormat = SaveFormat.Text;
            // All headers and footers are placed at the very end of the output document
            options.ExportHeadersFootersMode = TxtExportHeadersFootersMode.AllAtEnd;
            
            doc.Save(ArtifactsDir + "ExportHeadersFootersModeA.txt", options);

            // Only primary headers and footers are exported at the beginning and end of each section
            options.ExportHeadersFootersMode = TxtExportHeadersFootersMode.PrimaryOnly;
            
            doc.Save(ArtifactsDir + "ExportHeadersFootersModeB.txt", options);

            // No headers and footers are exported
            options.ExportHeadersFootersMode = TxtExportHeadersFootersMode.None;
            
            doc.Save(ArtifactsDir + "ExportHeadersFootersModeC.txt", options);
            //ExEnd:ExportHeadersFootersMode
        }

        [Test]
        public static void UseTabCharacterPerLevelForListIndentation()
        {
            //ExStart:UseTabCharacterPerLevelForListIndentation
            Document doc = new Document(LoadingSavingDir + "List indentation.docx");

            TxtSaveOptions options = new TxtSaveOptions();
            options.ListIndentation.Count = 1;
            options.ListIndentation.Character = '\t';

            doc.Save(ArtifactsDir + "UseTabCharacterPerLevelForListIndentation.txt", options);
            //ExEnd:UseTabCharacterPerLevelForListIndentation
        }

        [Test]
        public static void UseSpaceCharacterPerLevelForListIndentation()
        {
            //ExStart:UseSpaceCharacterPerLevelForListIndentation
            Document doc = new Document(LoadingSavingDir + "List indentation.docx");

            TxtSaveOptions options = new TxtSaveOptions();
            options.ListIndentation.Count = 3;
            options.ListIndentation.Character = ' ';

            doc.Save(ArtifactsDir + "UseSpaceCharacterPerLevelForListIndentation.txt", options);
            //ExEnd:UseSpaceCharacterPerLevelForListIndentation
        }

        [Test]
        public static void DefaultLevelForListIndentation()
        {
            //ExStart:DefaultLevelForListIndentation
            Document doc1 = new Document(LoadingSavingDir + "List indentation.docx");
            doc1.Save(ArtifactsDir + "DefaultLevelForListIndentation1.txt");

            Document doc2 = new Document(LoadingSavingDir + "List indentation.docx");
            TxtSaveOptions options = new TxtSaveOptions();
            doc2.Save(ArtifactsDir + "DefaultLevelForListIndentation2.txt", options);
            //ExEnd:DefaultLevelForListIndentation
        }
    }
}