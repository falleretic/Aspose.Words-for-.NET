using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Styles
{
    class CopyStyles : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:CopyStylesFromDocument
            Document doc = new Document();
            Document target = new Document(StyleDir + "Styles.docx");
            
            target.CopyStylesFromTemplate(doc);
            
            doc.Save(ArtifactsDir + "CopyStyles.docx");
            //ExEnd:CopyStylesFromDocument
        }
    }
}