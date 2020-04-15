using Aspose.Words.Fields;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Fields
{
    class InsertFieldNone : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:InsertFieldNone
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            FieldUnknown field = (FieldUnknown) builder.InsertField(FieldType.FieldNone, false);

            doc.Save(ArtifactsDir + "InsertFieldNone.docx");
            //ExEnd:InsertFieldNone
        }
    }
}