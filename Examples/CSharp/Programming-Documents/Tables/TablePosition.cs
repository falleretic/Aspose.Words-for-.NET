using System;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Tables
{
    class TablePosition : TestDataHelper
    {
        [Test]
        public static void GetTablePosition()
        {
            //ExStart:GetTablePosition
            Document doc = new Document(TablesDir + "Table.Document.doc");

            // Retrieve the first table in the document
            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);

            if (table.TextWrapping == TextWrapping.Around)
            {
                Console.WriteLine(table.RelativeHorizontalAlignment);
                Console.WriteLine(table.RelativeVerticalAlignment);
            }
            else
            {
                Console.WriteLine(table.Alignment);
            }
            //ExEnd:GetTablePosition
        }

        [Test]
        public static void GetFloatingTablePosition()
        {
            //ExStart:GetFloatingTablePosition
            Document doc = new Document(TablesDir + "FloatingTablePosition.docx");
            
            foreach (Table table in doc.FirstSection.Body.Tables)
            {
                // If table is floating type then print its positioning properties
                if (table.TextWrapping == TextWrapping.Around)
                {
                    Console.WriteLine(table.HorizontalAnchor);
                    Console.WriteLine(table.VerticalAnchor);
                    Console.WriteLine(table.AbsoluteHorizontalDistance);
                    Console.WriteLine(table.AbsoluteVerticalDistance);
                    Console.WriteLine(table.AllowOverlap);
                    Console.WriteLine(table.AbsoluteHorizontalDistance);
                    Console.WriteLine(table.RelativeVerticalAlignment);
                    Console.WriteLine("..............................");
                }
            }
            //ExEnd:GetFloatingTablePosition
        }

        [Test]
        public static void SetFloatingTablePosition()
        {
            //ExStart:SetFloatingTablePosition
            Document doc = new Document(TablesDir + "FloatingTablePosition.docx");

            Table table = doc.FirstSection.Body.Tables[0];
            // Sets absolute table horizontal position at 10pt
            table.AbsoluteHorizontalDistance = 10;
            // Sets vertical table position to center of entity specified by Table.VerticalAnchor
            table.RelativeVerticalAlignment = VerticalAlignment.Center;

            doc.Save(ArtifactsDir + "SetFloatingTablePosition.docx");
            //ExEnd:SetFloatingTablePosition
        }

        [Test]
        public static void SetRelativeHorizontalOrVerticalPosition()
        {
            // ExStart:SetRelativeHorizontalOrVerticalPosition
            Document doc = new Document(TablesDir + "FloatingTablePosition.docx");
            Table table = doc.FirstSection.Body.Tables[0];

            table.HorizontalAnchor = RelativeHorizontalPosition.Column;
            table.VerticalAnchor = RelativeVerticalPosition.Page;

            // Save the document to disk.
            doc.Save(ArtifactsDir + "Table.SetFloatingTablePosition.docx");
            // ExEnd:SetRelativeHorizontalOrVerticalPosition
        }
    }
}