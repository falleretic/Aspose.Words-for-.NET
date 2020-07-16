using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Sections
{
    class AppendSectionContent : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:AppendSectionContent
            Document doc = new Document(SectionsDir + "Section.AppendContent.doc");
            // This is the section that we will append and prepend to
            Section section = doc.Sections[2];

            // This copies content of the 1st section and inserts it at the beginning of the specified section
            Section sectionToPrepend = doc.Sections[0];
            section.PrependContent(sectionToPrepend);

            // This copies content of the 2nd section and inserts it at the end of the specified section
            Section sectionToAppend = doc.Sections[1];
            section.AppendContent(sectionToAppend);
            //ExEnd:AppendSectionContent
        }
    }
}