using System.Drawing;
using Aspose.Words.Lists;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_with_Documents.Document_Content
{
    class WorkingWithList : TestDataHelper
    {
        [Test]
        public static void SetRestartAtEachSection()
        {
            //ExStart:SetRestartAtEachSection
            Document doc = new Document();
            
            doc.Lists.Add(ListTemplate.NumberDefault);

            Lists.List list = doc.Lists[0];
            // Set true to specify that the list has to be restarted at each section
            list.IsRestartAtEachSection = true;

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.ListFormat.List = list;

            for (int i = 1; i < 45; i++)
            {
                builder.Writeln(string.Format("List Item {0}", i));

                // Insert section break
                if (i == 15)
                    builder.InsertBreak(BreakType.SectionBreakNewPage);
            }

            // IsRestartAtEachSection will be written only if compliance is higher then OoxmlComplianceCore.Ecma376
            OoxmlSaveOptions options = new OoxmlSaveOptions();
            options.Compliance = OoxmlCompliance.Iso29500_2008_Transitional;

            doc.Save(ArtifactsDir + "SetRestartAtEachSection.docx", options);
            //ExEnd:SetRestartAtEachSection
        }

        [Test]
        public static void SpecifyListLevel()
        {
            //ExStart:SpecifyListLevel
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a numbered list based on one of the Microsoft Word list templates and
            // apply it to the current paragraph in the document builder
            builder.ListFormat.List = doc.Lists.Add(ListTemplate.NumberArabicDot);

            // There are 9 levels in this list, lets try them all
            for (int i = 0; i < 9; i++)
            {
                builder.ListFormat.ListLevelNumber = i;
                builder.Writeln("Level " + i);
            }

            // Create a bulleted list based on one of the Microsoft Word list templates
            // and apply it to the current paragraph in the document builder
            builder.ListFormat.List = doc.Lists.Add(ListTemplate.BulletDiamonds);

            // There are 9 levels in this list, lets try them all
            for (int i = 0; i < 9; i++)
            {
                builder.ListFormat.ListLevelNumber = i;
                builder.Writeln("Level " + i);
            }

            // This is a way to stop list formatting
            builder.ListFormat.List = null;

            builder.Document.Save(ArtifactsDir + "SpecifyListLevel.docx");
            // ExEnd:SpecifyListLevel
        }

        [Test]
        public static void RestartListNumber()
        {
            //ExStart:RestartListNumber
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a list based on a template
            Lists.List list1 = doc.Lists.Add(ListTemplate.NumberArabicParenthesis);
            // Modify the formatting of the list
            list1.ListLevels[0].Font.Color = Color.Red;
            list1.ListLevels[0].Alignment = ListLevelAlignment.Right;

            builder.Writeln("List 1 starts below:");
            // Use the first list in the document for a while
            builder.ListFormat.List = list1;
            builder.Writeln("Item 1");
            builder.Writeln("Item 2");
            builder.ListFormat.RemoveNumbers();

            // Now I want to reuse the first list, but need to restart numbering
            // This should be done by creating a copy of the original list formatting
            Lists.List list2 = doc.Lists.AddCopy(list1);

            // We can modify the new list in any way. Including setting new start number
            list2.ListLevels[0].StartAt = 10;

            // Use the second list in the document
            builder.Writeln("List 2 starts below:");
            builder.ListFormat.List = list2;
            builder.Writeln("Item 1");
            builder.Writeln("Item 2");
            builder.ListFormat.RemoveNumbers();

            builder.Document.Save(ArtifactsDir + "RestartListNumber.docx");
            //ExEnd:RestartListNumber
        }
    }
}