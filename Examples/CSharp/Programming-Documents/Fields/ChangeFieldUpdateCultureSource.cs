using System;
using Aspose.Words.Fields;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Fields
{
    class ChangeFieldUpdateCultureSource : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:ChangeFieldUpdateCultureSource
            //ExStart:DocumentBuilderInsertField
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert content with German locale
            builder.Font.LocaleId = 1031;
            builder.InsertField("MERGEFIELD Date1 \\@ \"dddd, d MMMM yyyy\"");
            builder.Write(" - ");
            builder.InsertField("MERGEFIELD Date2 \\@ \"dddd, d MMMM yyyy\"");
            //ExEnd:DocumentBuilderInsertField

            // Shows how to specify where the culture used for date formatting during field update and mail merge is chosen from
            // Set the culture used during field update to the culture used by the field
            doc.FieldOptions.FieldUpdateCultureSource = FieldUpdateCultureSource.FieldCode;
            doc.MailMerge.Execute(new string[] { "Date2" }, new object[] { new DateTime(2011, 1, 01) });
            
            doc.Save(ArtifactsDir + "Field.ChangeFieldUpdateCultureSource.doc");
            //ExEnd:ChangeFieldUpdateCultureSource
        }
    }
}