using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Fields
{
    class InsertField : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:InsertField
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.InsertField(@"MERGEFIELD MyFieldName \* MERGEFORMAT");
            
            doc.Save(ArtifactsDir + "InsertField.docx");
            //ExEnd:InsertField
        }
    }
}