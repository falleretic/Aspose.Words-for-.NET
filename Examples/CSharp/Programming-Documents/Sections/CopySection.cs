using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Sections
{
    class CopySection : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:CopySection
            Document srcDoc = new Document(SectionsDir + "Document.doc");
            Document dstDoc = new Document();

            Section sourceSection = srcDoc.Sections[0];
            Section newSection = (Section) dstDoc.ImportNode(sourceSection, true);
            dstDoc.Sections.Add(newSection);
            
            dstDoc.Save(ArtifactsDir + "CopySection.docx");
            //ExEnd:CopySection
        }
    }
}