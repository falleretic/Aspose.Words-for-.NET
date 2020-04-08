namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Sections
{
    class CloneSection : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:CloneSection
            Document doc = new Document(SectionsDir + "Document.doc");
            Section cloneSection = doc.Sections[0].Clone();
            //ExEnd:CloneSection
        }
    }
}