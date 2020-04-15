using Aspose.Words.Markup;
using Aspose.Words.Tables;
using System;
using System.Drawing;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_StructuredDocumentTag
{
    internal class WorkingWithSdt : TestDataHelper
    {
        [Test]
        public static void SetContentControlColor()
        {
            //ExStart:SetContentControlColor
            Document doc = new Document(SdtDir + "input.docx");
            StructuredDocumentTag sdt = (StructuredDocumentTag) doc.GetChild(NodeType.StructuredDocumentTag, 0, true);
            sdt.Color = Color.Red;

            doc.Save(ArtifactsDir + "SetContentControlColor.docx");
            //ExEnd:SetContentControlColor
        }

        [Test]
        public static void ClearContentsControl()
        {
            //ExStart:ClearContentsControl
            Document doc = new Document(SdtDir + "input.docx");
            StructuredDocumentTag sdt = (StructuredDocumentTag) doc.GetChild(NodeType.StructuredDocumentTag, 0, true);
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

            StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Block);
            doc.FirstSection.Body.AppendChild(sdt);

            sdt.XmlMapping.SetMapping(xmlPart, "/root[1]/text[1]", "");

            doc.Save(ArtifactsDir + "BindSDTtoCustomXmlPart.doc");
            //ExEnd:BindSDTtoCustomXmlPart
        }

        [Test]
        public static void SetContentControlStyle()
        {
            //ExStart:SetContentControlStyle
            Document doc = new Document(SdtDir + "input.docx");
            StructuredDocumentTag sdt = (StructuredDocumentTag) doc.GetChild(NodeType.StructuredDocumentTag, 0, true);
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

            StructuredDocumentTag repeatingSectionSdt =
                new StructuredDocumentTag(doc, SdtType.RepeatingSection, MarkupLevel.Row);
            repeatingSectionSdt.XmlMapping.SetMapping(xmlPart, "/books[1]/book", "");
            table.AppendChild(repeatingSectionSdt);

            StructuredDocumentTag repeatingSectionItemSdt =
                new StructuredDocumentTag(doc, SdtType.RepeatingSectionItem, MarkupLevel.Row);
            repeatingSectionSdt.AppendChild(repeatingSectionItemSdt);

            Row row = new Row(doc);
            repeatingSectionItemSdt.AppendChild(row);

            StructuredDocumentTag titleSdt =
                new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Cell);
            titleSdt.XmlMapping.SetMapping(xmlPart, "/books[1]/book[1]/title[1]", "");
            row.AppendChild(titleSdt);

            StructuredDocumentTag authorSdt =
                new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Cell);
            authorSdt.XmlMapping.SetMapping(xmlPart, "/books[1]/book[1]/author[1]", "");
            row.AppendChild(authorSdt);

            doc.Save(ArtifactsDir + "Document.docx");
            //ExEnd:CreatingTableRepeatingSectionMappedToCustomXmlPart
        }
    }
}