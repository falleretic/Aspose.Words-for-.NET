using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Sections
{
    class ModifyPageSetupInAllSectionsOfDocument : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:ModifyPageSetupInAllSectionsOfDocument
            Document doc = new Document(SectionsDir + "ModifyPageSetupInAllSections.doc");

            // It is important to understand that a document can contain many sections and each
            // section has its own page setup. In this case we want to modify them all
            foreach (Section section in doc)
                section.PageSetup.PaperSize = PaperSize.Letter;

            doc.Save(ArtifactsDir + "ModifyPageSetupInAllSections.doc");
            //ExEnd:ModifyPageSetupInAllSectionsOfDocument
        }
    }
}