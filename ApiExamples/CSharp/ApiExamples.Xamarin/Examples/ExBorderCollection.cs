﻿// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.Collections.Generic;
using System.Drawing;
using Aspose.Words;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExBorderCollection : ApiExampleBase
    {
        [Test]
        public void GetBordersEnumerator()
        {
            //ExStart
            //ExFor:BorderCollection.GetEnumerator
            //ExSummary:Shows how to enumerate all borders in a collection.
            Document doc = new Document(MyDir + "Borders.docx");
            DocumentBuilder builder = new DocumentBuilder(doc);

            BorderCollection borders = builder.ParagraphFormat.Borders;

            using (IEnumerator<Border> enumerator = borders.GetEnumerator())
            {
                while (enumerator.MoveNext())
                {
                    // Do something useful.
                    Border b = enumerator.Current;
                    b.Color = Color.RoyalBlue;
                    b.LineStyle = LineStyle.Double;
                }
            }

            doc.Save(ArtifactsDir + "BorderCollection.GetBordersEnumerator.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "BorderCollection.GetBordersEnumerator.docx");

            foreach (Border border in doc.FirstSection.Body.FirstParagraph.ParagraphFormat.Borders)
            {
                Assert.AreEqual(Color.RoyalBlue.ToArgb(), border.Color.ToArgb());
                Assert.AreEqual(LineStyle.Double, border.LineStyle);
            }
        }

        [Test]
        public void RemoveAllBorders()
        {
            //ExStart
            //ExFor:BorderCollection.ClearFormatting
            //ExSummary:Shows how to remove all borders from a paragraph at once.
            Document doc = new Document(MyDir + "Borders.docx");
            DocumentBuilder builder = new DocumentBuilder(doc);
            BorderCollection borders = builder.ParagraphFormat.Borders;

            borders.ClearFormatting();

            doc.Save(ArtifactsDir + "BorderCollection.RemoveAllBorders.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "BorderCollection.RemoveAllBorders.docx");

            foreach (Border border in doc.FirstSection.Body.FirstParagraph.ParagraphFormat.Borders)
            {
                Assert.AreEqual(Color.Empty.ToArgb(), border.Color.ToArgb());
                Assert.AreEqual(LineStyle.None, border.LineStyle);
            }
        }
    }
}