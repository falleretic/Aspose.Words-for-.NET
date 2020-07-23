using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Sections
{
    class ModifyPageSetupInAllSectionsOfDocument : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:ModifyPageSetupInAllSectionsOfDocument
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello1");
            doc.AppendChild(new Section(doc));
            builder.Writeln("Hello22");
            doc.AppendChild(new Section(doc));
            builder.Writeln("Hello3");
            doc.AppendChild(new Section(doc));
            builder.Writeln("Hello45");

            // It is important to understand that a document can contain many sections and each
            // section has its own page setup. In this case we want to modify them all
            foreach (Section section in doc)
                section.PageSetup.PaperSize = PaperSize.Letter;

            doc.Save(ArtifactsDir + "ModifyPageSetupInAllSections.doc");
            //ExEnd:ModifyPageSetupInAllSectionsOfDocument
        }
    }
}