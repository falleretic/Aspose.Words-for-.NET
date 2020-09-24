﻿using System;
using System.Drawing;
using Aspose.Words.Drawing;
using Aspose.Words.Markup;
using Aspose.Words.Tables;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.DocumentEx
{
    class WorkingWithSDT : TestDataHelper
    {
        [Test]
        public static void CheckBoxTypeContentControl()
        {
            //ExStart:CheckBoxTypeContentControl
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Markup.StructuredDocumentTag sdtCheckBox = new Markup.StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline);
            // Insert content control into the document
            builder.InsertNode(sdtCheckBox);
            
            doc.Save(ArtifactsDir + "CheckBoxTypeContentControl.docx", SaveFormat.Docx);
            //ExEnd:CheckBoxTypeContentControl
        }

        [Test]
        public static void SetCurrentStateOfCheckBox()
        {
            //ExStart:SetCurrentStateOfCheckBox
            // Open an existing document
            Document doc = new Document(DocumentDir + "Structured document tags.docx");
            
            // Get the first content control from the document
            Markup.StructuredDocumentTag sdtCheckBox =
                (Markup.StructuredDocumentTag) doc.GetChild(NodeType.StructuredDocumentTag, 0, true);

            // StructuredDocumentTag.Checked property gets/sets current state of the Checkbox SDT
            if (sdtCheckBox.SdtType == SdtType.Checkbox)
                sdtCheckBox.Checked = true;

            doc.Save(ArtifactsDir + "SetCurrentStateOfCheckBox.docx");
            //ExEnd:SetCurrentStateOfCheckBox
        }

        [Test]
        public static void ModifyContentControls()
        {
            //ExStart:ModifyContentControls
            // Open an existing document
            Document doc = new Document(DocumentDir + "Structured document tags.docx");

            foreach (Markup.StructuredDocumentTag sdt in doc.GetChildNodes(NodeType.StructuredDocumentTag, true))
            {
                switch (sdt.SdtType)
                {
                    case SdtType.PlainText:
                    {
                        sdt.RemoveAllChildren();
                        Paragraph para = sdt.AppendChild(new Paragraph(doc)) as Paragraph;
                        Run run = new Run(doc, "new text goes here");
                        para.AppendChild(run);
                        break;
                    }
                    case SdtType.DropDownList:
                    {
                        SdtListItem secondItem = sdt.ListItems[2];
                        sdt.ListItems.SelectedValue = secondItem;
                        break;
                    }
                    case SdtType.Picture:
                    {
                        Shape shape = (Shape) sdt.GetChild(NodeType.Shape, 0, true);
                        if (shape.HasImage)
                        {
                            shape.ImageData.SetImage(DocumentDir + "Watermark.png");
                        }

                        break;
                    }
                }
            }
            
            doc.Save(ArtifactsDir + "ModifyContentControls.docx");
            //ExEnd:ModifyContentControls
        }

        [Test]
        public static void ComboBoxContentControl()
        {
            //ExStart:ComboBoxContentControl
            Document doc = new Document();
            Markup.StructuredDocumentTag sdt = new Markup.StructuredDocumentTag(doc, SdtType.ComboBox, MarkupLevel.Block);

            sdt.ListItems.Add(new SdtListItem("Choose an item", "-1"));
            sdt.ListItems.Add(new SdtListItem("Item 1", "1"));
            sdt.ListItems.Add(new SdtListItem("Item 2", "2"));
            doc.FirstSection.Body.AppendChild(sdt);

            doc.Save(ArtifactsDir + "ComboBoxContentControl.docx");
            //ExEnd:ComboBoxContentControl
        }

        [Test]
        public static void RichTextBoxContentControl()
        {
            //ExStart:RichTextBoxContentControl
            Document doc = new Document();
            Markup.StructuredDocumentTag sdtRichText = new Markup.StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block);

            Paragraph para = new Paragraph(doc);
            Run run = new Run(doc);
            run.Text = "Hello World";
            run.Font.Color = Color.Green;
            para.Runs.Add(run);
            sdtRichText.ChildNodes.Add(para);
            doc.FirstSection.Body.AppendChild(sdtRichText);

            doc.Save(ArtifactsDir + "RichTextBoxContentControl.docx");
            //ExEnd:RichTextBoxContentControl
        }

        [Test]
        public static void SetContentControlColor()
        {
            //ExStart:SetContentControlColor
            Document doc = new Document(SdtDir + "Structured document tags.docx");
            Markup.StructuredDocumentTag sdt = (Markup.StructuredDocumentTag) doc.GetChild(NodeType.StructuredDocumentTag, 0, true);
            sdt.Color = Color.Red;

            doc.Save(ArtifactsDir + "SetContentControlColor.docx");
            //ExEnd:SetContentControlColor
        }

        [Test]
        public static void ClearContentsControl()
        {
            //ExStart:ClearContentsControl
            Document doc = new Document(SdtDir + "Structured document tags.docx");
            Markup.StructuredDocumentTag sdt = (Markup.StructuredDocumentTag) doc.GetChild(NodeType.StructuredDocumentTag, 0, true);
            sdt.Clear();

            doc.Save(ArtifactsDir + "ClearContentsControl.doc");
            //ExEnd:ClearContentsControl
        }

        [Test]
        public static void BindSdTtoCustomXmlPart()
        {
            //ExStart:BindSDTtoCustomXmlPart
            Document doc = new Document();
            CustomXmlPart xmlPart =
                doc.CustomXmlParts.Add(Guid.NewGuid().ToString("B"), "<root><text>Hello, World!</text></root>");

            Markup.StructuredDocumentTag sdt = new Markup.StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Block);
            doc.FirstSection.Body.AppendChild(sdt);

            sdt.XmlMapping.SetMapping(xmlPart, "/root[1]/text[1]", "");

            doc.Save(ArtifactsDir + "BindSDTtoCustomXmlPart.doc");
            //ExEnd:BindSDTtoCustomXmlPart
        }

        [Test]
        public static void SetContentControlStyle()
        {
            //ExStart:SetContentControlStyle
            Document doc = new Document(SdtDir + "Structured document tags.docx");
            Markup.StructuredDocumentTag sdt = (Markup.StructuredDocumentTag) doc.GetChild(NodeType.StructuredDocumentTag, 0, true);
            Style style = doc.Styles[StyleIdentifier.Quote];
            sdt.Style = style;

            doc.Save(ArtifactsDir + "SetContentControlStyle.docx");
            //ExEnd:SetContentControlStyle
        }

        [Test]
        public static void CreatingTableRepeatingSectionMappedToCustomXmlPart()
        {
            //ExStart:CreatingTableRepeatingSectionMappedToCustomXmlPart
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            CustomXmlPart xmlPart = doc.CustomXmlParts.Add("Books",
                "<books><book><title>Everyday Italian</title><author>Giada De Laurentiis</author></book>" +
                "<book><title>Harry Potter</title><author>J K. Rowling</author></book>" +
                "<book><title>Learning XML</title><author>Erik T. Ray</author></book></books>");

            Table table = builder.StartTable();

            builder.InsertCell();
            builder.Write("Title");

            builder.InsertCell();
            builder.Write("Author");

            builder.EndRow();
            builder.EndTable();

            Markup.StructuredDocumentTag repeatingSectionSdt =
                new Markup.StructuredDocumentTag(doc, SdtType.RepeatingSection, MarkupLevel.Row);
            repeatingSectionSdt.XmlMapping.SetMapping(xmlPart, "/books[1]/book", "");
            table.AppendChild(repeatingSectionSdt);

            Markup.StructuredDocumentTag repeatingSectionItemSdt =
                new Markup.StructuredDocumentTag(doc, SdtType.RepeatingSectionItem, MarkupLevel.Row);
            repeatingSectionSdt.AppendChild(repeatingSectionItemSdt);

            Row row = new Row(doc);
            repeatingSectionItemSdt.AppendChild(row);

            Markup.StructuredDocumentTag titleSdt =
                new Markup.StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Cell);
            titleSdt.XmlMapping.SetMapping(xmlPart, "/books[1]/book[1]/title[1]", "");
            row.AppendChild(titleSdt);

            Markup.StructuredDocumentTag authorSdt =
                new Markup.StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Cell);
            authorSdt.XmlMapping.SetMapping(xmlPart, "/books[1]/book[1]/author[1]", "");
            row.AppendChild(authorSdt);

            doc.Save(ArtifactsDir + "Document.docx");
            //ExEnd:CreatingTableRepeatingSectionMappedToCustomXmlPart
        }
    }
}