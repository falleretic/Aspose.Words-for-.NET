﻿using System;
using System.Drawing;
using Aspose.Words.Tables;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Tables
{
    class ApplyStyle : TestDataHelper
    {
        /// <summary>
        /// Shows how to build a new table with a table style applied.
        /// </summary>
        [Test]
        public static void BuildTableWithStyle()
        {
            //ExStart:BuildTableWithStyle
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Table table = builder.StartTable();
            // We must insert at least one row first before setting any table formatting
            builder.InsertCell();
            // Set the table style used based of the unique style identifier
            // Note that not all table styles are available when saving as .doc format
            table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;
            // Apply which features should be formatted by the style
            table.StyleOptions =
                TableStyleOptions.FirstColumn | TableStyleOptions.RowBands | TableStyleOptions.FirstRow;
            table.AutoFit(AutoFitBehavior.AutoFitToContents);

            // Continue with building the table as normal
            builder.Writeln("Item");
            builder.CellFormat.RightPadding = 40;
            builder.InsertCell();
            builder.Writeln("Quantity (kg)");
            builder.EndRow();

            builder.InsertCell();
            builder.Writeln("Apples");
            builder.InsertCell();
            builder.Writeln("20");
            builder.EndRow();

            builder.InsertCell();
            builder.Writeln("Bananas");
            builder.InsertCell();
            builder.Writeln("40");
            builder.EndRow();

            builder.InsertCell();
            builder.Writeln("Carrots");
            builder.InsertCell();
            builder.Writeln("50");
            builder.EndRow();

            doc.Save(ArtifactsDir + "BuildTableWithStyle.docx");
            //ExEnd:BuildTableWithStyle
        }

        /// <summary>
        /// Shows how to expand the formatting from styles onto the rows and cells of the table as direct formatting.
        /// </summary>
        [Test]
        public static void ExpandFormattingOnCellsAndRowFromStyle()
        {
            //ExStart:ExpandFormattingOnCellsAndRowFromStyle
            Document doc = new Document(TablesDir + "Tables.docx");

            // Get the first cell of the first table in the document
            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);
            Cell firstCell = table.FirstRow.FirstCell;

            // First print the color of the cell shading
            // This should be empty as the current shading is stored in the table style
            Color cellShadingBefore = firstCell.CellFormat.Shading.BackgroundPatternColor;
            Console.WriteLine("Cell shading before style expansion: " + cellShadingBefore);

            // Expand table style formatting to direct formatting
            doc.ExpandTableStylesToDirectFormatting();

            // Now print the cell shading after expanding table styles
            // A blue background pattern color should have been applied from the table style
            Color cellShadingAfter = firstCell.CellFormat.Shading.BackgroundPatternColor;
            Console.WriteLine("Cell shading after style expansion: " + cellShadingAfter);
            //ExEnd:ExpandFormattingOnCellsAndRowFromStyle
        }

        [Test]
        public static void CreateTableStyle()
        {
            //ExStart:CreateTableStyle
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Name");
            builder.InsertCell();
            builder.Write("Value");
            builder.EndRow();
            builder.InsertCell();
            builder.InsertCell();
            builder.EndTable();

            TableStyle tableStyle = (TableStyle) doc.Styles.Add(StyleType.Table, "MyTableStyle1");
            tableStyle.Borders.LineStyle = LineStyle.Double;
            tableStyle.Borders.LineWidth = 1;
            tableStyle.LeftPadding = 18;
            tableStyle.RightPadding = 18;
            tableStyle.TopPadding = 12;
            tableStyle.BottomPadding = 12;

            table.Style = tableStyle;

            doc.Save(ArtifactsDir + "CreateTableStyle.docx");
            //ExEnd:CreateTableStyle
        }

        [Test]
        public static void DefineConditionalFormatting()
        {
            //ExStart:DefineConditionalFormatting
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Name");
            builder.InsertCell();
            builder.Write("Value");
            builder.EndRow();
            builder.InsertCell();
            builder.InsertCell();
            builder.EndTable();

            TableStyle tableStyle = (TableStyle) doc.Styles.Add(StyleType.Table, "MyTableStyle1");
            // Define background color to the first row of table
            tableStyle.ConditionalStyles.FirstRow.Shading.BackgroundPatternColor = Color.GreenYellow;
            tableStyle.ConditionalStyles.FirstRow.Shading.Texture = TextureIndex.TextureNone;

            table.Style = tableStyle;

            doc.Save(ArtifactsDir + "DefineConditionalFormatting.docx");
            //ExEnd:DefineConditionalFormatting
        }
    }
}