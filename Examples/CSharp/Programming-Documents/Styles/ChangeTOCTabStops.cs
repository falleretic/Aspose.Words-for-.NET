namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Styles
{
    class ChangeTocTabStops : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:ChangeTOCTabStops
            Document doc = new Document(StyleDir + "Document.TableOfContents.doc");

            // Iterate through all paragraphs in the document
            foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
            {
                // Check if this paragraph is formatted using the TOC result based styles
                // This is any style between TOC and TOC9
                if (para.ParagraphFormat.Style.StyleIdentifier >= StyleIdentifier.Toc1 &&
                    para.ParagraphFormat.Style.StyleIdentifier <= StyleIdentifier.Toc9)
                {
                    // Get the first tab used in this paragraph, this should be the tab used to align the page numbers
                    TabStop tab = para.ParagraphFormat.TabStops[0];
                    // Remove the old tab from the collection
                    para.ParagraphFormat.TabStops.RemoveByPosition(tab.Position);
                    // Insert a new tab using the same properties but at a modified position
                    // We could also change the separators used (dots) by passing a different Leader type
                    para.ParagraphFormat.TabStops.Add(tab.Position - 50, tab.Alignment, tab.Leader);
                }
            }

            doc.Save(ArtifactsDir + "Document.TableOfContents.doc");
            //ExEnd:ChangeTOCTabStops
        }
    }
}