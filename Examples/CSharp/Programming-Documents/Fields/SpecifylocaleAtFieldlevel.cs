using Aspose.Words.Fields;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Fields
{
    class SpecifyLocaleAtFieldLevel : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:SpecifylocaleAtFieldlevel
            DocumentBuilder builder = new DocumentBuilder();

            Field field = builder.InsertField(FieldType.FieldDate, true);
            field.LocaleId = 1049;
            
            builder.Document.Save(ArtifactsDir + "SpecifylocaleAtFieldlevel.docx");
            //ExEnd:SpecifylocaleAtFieldlevel
        }
    }
}