﻿using System.Drawing;
using System.IO;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.DocumentEx
{
    class DocumentBuilderInsertElements : TestDataHelper
    {
        [Test]
        public static void InsertTextInputFormField()
        {
            //ExStart:DocumentBuilderInsertTextInputFormField
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertTextInput("TextInput", TextFormFieldType.Regular, "", "Hello", 0);
            
            doc.Save(ArtifactsDir + "DocumentBuilderInsertTextInputFormField.doc");
            //ExEnd:DocumentBuilderInsertTextInputFormField
        }

        [Test]
        public static void InsertCheckBoxFormField()
        {
            //ExStart:DocumentBuilderInsertCheckBoxFormField
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertCheckBox("CheckBox", true, true, 0);
            
            doc.Save(ArtifactsDir + "DocumentBuilderInsertCheckBoxFormField.doc");
            //ExEnd:DocumentBuilderInsertCheckBoxFormField
        }

        [Test]
        public static void InsertComboBoxFormField()
        {
            //ExStart:DocumentBuilderInsertComboBoxFormField
            string[] items = { "One", "Two", "Three" };
            
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertComboBox("DropDown", items, 0);
            
            doc.Save(ArtifactsDir + "DocumentBuilderInsertComboBoxFormField.doc");
            //ExEnd:DocumentBuilderInsertComboBoxFormField
        }

        [Test]
        public static void InsertHtml()
        {
            //ExStart:DocumentBuilderInsertHtml
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertHtml(
                "<P align='right'>Paragraph right</P>" +
                "<b>Implicit paragraph left</b>" +
                "<div align='center'>Div center</div>" +
                "<h1 align='left'>Heading 1 left.</h1>");
            
            doc.Save(ArtifactsDir + "DocumentBuilderInsertHtml.doc");
            //ExEnd:DocumentBuilderInsertHtml
        }

        [Test]
        public static void InsertHyperlink()
        {
            //ExStart:DocumentBuilderInsertHyperlink
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Write("Please make sure to visit ");
            // Specify font formatting for the hyperlink
            builder.Font.Color = Color.Blue;
            builder.Font.Underline = Underline.Single;
            // Insert the link
            builder.InsertHyperlink("Aspose Website", "http://www.aspose.com", false);
            // Revert to default formatting
            builder.Font.ClearFormatting();
            builder.Write(" for more information.");

            doc.Save(ArtifactsDir + "DocumentBuilderInsertHyperlink.doc");
            //ExEnd:DocumentBuilderInsertHyperlink
        }

        [Test]
        public static void InsertTableOfContents()
        {
            //ExStart:DocumentBuilderInsertTableOfContents
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            // Insert a table of contents at the beginning of the document
            builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
            // Start the actual document content on the second page
            builder.InsertBreak(BreakType.PageBreak);
            
            // Build a document with complex structure by applying different heading styles thus creating TOC entries
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;

            builder.Writeln("Heading 1");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;

            builder.Writeln("Heading 1.1");
            builder.Writeln("Heading 1.2");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;

            builder.Writeln("Heading 2");
            builder.Writeln("Heading 3");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;

            builder.Writeln("Heading 3.1");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;

            builder.Writeln("Heading 3.1.1");
            builder.Writeln("Heading 3.1.2");
            builder.Writeln("Heading 3.1.3");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;

            builder.Writeln("Heading 3.2");
            builder.Writeln("Heading 3.3");

            // The newly inserted table of contents will be initially empty
            // It needs to be populated by updating the fields in the document
            //ExStart:UpdateFields
            doc.UpdateFields();
            //ExEnd:UpdateFields

            doc.Save(ArtifactsDir + "DocumentBuilderInsertTableOfContents.doc");
            //ExEnd:DocumentBuilderInsertTableOfContents
        }

        [Test]
        public static void InsertOleObject()
        {
            //ExStart:DocumentBuilderInsertOleObject
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertOleObject("http://www.aspose.com", "htmlfile", true, true, null);
            
            doc.Save(ArtifactsDir + "DocumentBuilderInsertOleObject.doc");
            //ExEnd:DocumentBuilderInsertOleObject
        }

        [Test]
        public static void InsertOleObjectWithOlePackage()
        {
            //ExStart:InsertOleObjectwithOlePackage
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            byte[] bs = File.ReadAllBytes(DocumentDir + "Zip file.zip");
            using (Stream stream = new MemoryStream(bs))
            {
                Shape shape = builder.InsertOleObject(stream, "Package", true, null);
                OlePackage olePackage = shape.OleFormat.OlePackage;
                olePackage.FileName = "filename.zip";
                olePackage.DisplayName = "displayname.zip";
                
                doc.Save(ArtifactsDir + "DocumentBuilderInsertOleObjectOlePackage.doc");
            }
            //ExEnd:InsertOleObjectwithOlePackage

            //ExStart:GetAccessToOLEObjectRawData
            Shape oleShape = (Shape) doc.GetChild(NodeType.Shape, 0, true);
            byte[] oleRawData = oleShape.OleFormat.GetRawData();
            //ExEnd:GetAccessToOLEObjectRawData
        }
    }
}