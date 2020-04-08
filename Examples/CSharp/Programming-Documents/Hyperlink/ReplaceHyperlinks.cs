using Aspose.Words.Fields;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Hyperlink
{
    class ReplaceHyperlinks : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:ReplaceHyperlinks
            Document doc = new Document(HyperlinkDir + "ReplaceHyperlinks.doc");

            // Hyperlinks in a Word documents are fields
            foreach (Field field in doc.Range.Fields)
            {
                if (field.Type == FieldType.FieldHyperlink)
                {
                    FieldHyperlink hyperlink = (FieldHyperlink) field;

                    // Some hyperlinks can be local (links to bookmarks inside the document), ignore these
                    if (hyperlink.SubAddress != null)
                        continue;

                    hyperlink.Address = "http://www.aspose.com";
                    hyperlink.Result = "Aspose - The .NET & Java Component Publisher";
                }
            }

            doc.Save(ArtifactsDir + "ReplaceHyperlinks.doc");
            //ExEnd:ReplaceHyperlinks
        }
    }
}