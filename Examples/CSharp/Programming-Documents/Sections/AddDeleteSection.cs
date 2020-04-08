namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Sections
{
    class AddDeleteSection : TestDataHelper
    {
        public static void Run()
        {
            AddSection();
            DeleteSection();
            DeleteAllSections();
        }

        /// <summary>
        /// Shows how to add a section to the end of the document.
        /// </summary>
        private static void AddSection()
        {
            //ExStart:AddSection
            Document doc = new Document(SectionsDir + "Section.AddRemove.doc");
            Section sectionToAdd = new Section(doc);
            doc.Sections.Add(sectionToAdd);
            //ExEnd:AddSection
        }

        /// <summary>
        /// Shows how to remove a section at the specified index.
        /// </summary>
        private static void DeleteSection()
        {
            //ExStart:DeleteSection
            Document doc = new Document(SectionsDir + "Section.AddRemove.doc");
            doc.Sections.RemoveAt(0);
            //ExEnd:DeleteSection
        }

        /// <summary>
        /// Shows how to remove all sections from a document.
        /// </summary>
        private static void DeleteAllSections()
        {
            //ExStart:DeleteAllSections
            Document doc = new Document(SectionsDir + "Section.AddRemove.doc");
            doc.Sections.Clear();
            //ExEnd:DeleteAllSections
        }
    }
}