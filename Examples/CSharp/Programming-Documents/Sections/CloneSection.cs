using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Sections
{
    class CloneSection : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:CloneSection
            Document doc = new Document(SectionsDir + "Document.doc");
            Section cloneSection = doc.Sections[0].Clone();
            //ExEnd:CloneSection
        }
    }
}