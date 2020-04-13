using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using System;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Tables
{
    class TablePosition : TestDataHelper
    {
        public static void Run()
        {
            GetTablePosition();
            GetFloatingTablePosition();
            SetFloatingTablePosition();
        }

        private static void GetTablePosition()
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

        private static void GetFloatingTablePosition()
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

        private static void SetFloatingTablePosition()
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
    }
}