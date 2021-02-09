﻿// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Ole;
using Aspose.Words.Math;
using Aspose.Words.Rendering;
using Aspose.Words.Saving;
using Aspose.Words.Settings;
using NUnit.Framework;
using Color = System.Drawing.Color;
using DashStyle = Aspose.Words.Drawing.DashStyle;
using HorizontalAlignment = Aspose.Words.Drawing.HorizontalAlignment;
using TextBox = Aspose.Words.Drawing.TextBox;
#if NETCOREAPP2_1 || __MOBILE__
using SkiaSharp;
#endif

namespace ApiExamples
{
    /// <summary>
    /// Examples using shapes in documents.
    /// </summary>
    [TestFixture]
    public class ExShape : ApiExampleBase
    {
#if NET462 || JAVA
        [Test]
        public void AltText()
        {
            //ExStart
            //ExFor:ShapeBase.AlternativeText
            //ExFor:ShapeBase.Name
            //ExSummary:Shows how to use a shape's alternative text.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            Shape shape = builder.InsertShape(ShapeType.Cube, 150, 150);
            shape.Name = "MyCube";

            shape.AlternativeText = "Alt text for MyCube.";

            // We can access the alternative text of a shape by right-clicking it, and then via "Format AutoShape" -> "Alt Text".
            doc.Save(ArtifactsDir + "Shape.AltText.docx");

            // Save the document to HTML, and then delete the linked image that belongs to our shape.
            // The browser that is reading our HTML will display the alt text in place of the missing image.
            doc.Save(ArtifactsDir + "Shape.AltText.html");
            Assert.True(File.Exists(ArtifactsDir + "Shape.AltText.001.png")); //ExSkip
            File.Delete(ArtifactsDir + "Shape.AltText.001.png");
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.AltText.docx");
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyShape(ShapeType.Cube, "MyCube", 150.0d, 150.0d, 0, 0, shape);
            Assert.AreEqual("Alt text for MyCube.", shape.AlternativeText);
            Assert.AreEqual("Times New Roman", shape.Font.Name);

            doc = new Document(ArtifactsDir + "Shape.AltText.html");
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyShape(ShapeType.Image, string.Empty, 153.0d, 153.0d, 0, 0, shape);
            Assert.AreEqual("Alt text for MyCube.", shape.AlternativeText);

            TestUtil.FileContainsString(
                "<img src=\"Shape.AltText.001.png\" width=\"204\" height=\"204\" alt=\"Alt text for MyCube.\" " +
                "style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />", 
                ArtifactsDir + "Shape.AltText.html");
        }

        [TestCase(false)]
        [TestCase(true)]
        public void Font(bool hideShape)
        {
            //ExStart
            //ExFor:ShapeBase.Font
            //ExFor:ShapeBase.ParentParagraph
            //ExSummary:Shows how to insert a text box, and set the font of its contents.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello world!");

            Shape shape = builder.InsertShape(ShapeType.TextBox, 300, 50);
            builder.MoveTo(shape.LastParagraph);
            builder.Write("This text is inside the text box.");

            // Set the "Hidden" property of the shape's "Font" object to "true" to hide the text box from sight,
            // and collapse the space that it would normally occupy.
            // Set the "Hidden" property of the shape's "Font" object to "false" to leave the text box visible.
            shape.Font.Hidden = hideShape;

            // If the shape is visible, we will modify its appearance via the font object.
            if (!hideShape)
            {
                shape.Font.HighlightColor = Color.LightGray;
                shape.Font.Color = Color.Red;
                shape.Font.Underline = Underline.Dash;
            }
            
            // Move the builder out of the text box back into the main document.
            builder.MoveTo(shape.ParentParagraph);

            builder.Writeln("\nThis text is outside the text box.");

            doc.Save(ArtifactsDir + "Shape.Font.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.Font.docx");
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            Assert.AreEqual(hideShape, shape.Font.Hidden);

            if (hideShape)
            {
                Assert.AreEqual(Color.Empty.ToArgb(), shape.Font.HighlightColor.ToArgb());
                Assert.AreEqual(Color.Empty.ToArgb(), shape.Font.Color.ToArgb());
                Assert.AreEqual(Underline.None, shape.Font.Underline);
            }
            else
            {
                Assert.AreEqual(Color.Silver.ToArgb(), shape.Font.HighlightColor.ToArgb());
                Assert.AreEqual(Color.Red.ToArgb(), shape.Font.Color.ToArgb());
                Assert.AreEqual(Underline.Dash, shape.Font.Underline);
            }

            TestUtil.VerifyShape(ShapeType.TextBox, "TextBox 100002", 300.0d, 50.0d, 0, 0, shape);
            Assert.AreEqual("This text is inside the text box.", shape.GetText().Trim());
            Assert.AreEqual("Hello world!\rThis text is inside the text box.\r\rThis text is outside the text box.", doc.GetText().Trim());
        }

        [Test]
        public void Rotate()
        {
            //ExStart
            //ExFor:ShapeBase.CanHaveImage
            //ExFor:ShapeBase.Rotation
            //ExSummary:Shows how to insert and rotate an image.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a shape with an image.
            Shape shape = builder.InsertImage(Image.FromFile(ImageDir + "Logo.jpg"));
            Assert.True(shape.CanHaveImage);
            Assert.True(shape.HasImage);

            // Rotate the image 45 degrees clockwise.
            shape.Rotation = 45;

            doc.Save(ArtifactsDir + "Shape.Rotate.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.Rotate.docx");
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyShape(ShapeType.Image, string.Empty, 300.0d, 300.0d, 0, 0, shape);
            Assert.True(shape.CanHaveImage);
            Assert.True(shape.HasImage);
            Assert.AreEqual(45.0d, shape.Rotation);
        }

        [Test]
        public void AspectRatioLockedDefaultValue()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            Image image = Image.FromFile(ImageDir + "Transparent background logo.png");

            Shape shape = builder.InsertImage(image);
            shape.WrapType = WrapType.None;
            shape.BehindText = true;

            shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;

            // Calculate image left and top position so it appears in the center of the page.
            shape.Left = (builder.PageSetup.PageWidth - shape.Width) / 2;
            shape.Top = (builder.PageSetup.PageHeight - shape.Height) / 2;

            doc = DocumentHelper.SaveOpen(doc);

            shape = (Shape) doc.GetChild(NodeType.Shape, 0, true);
            Assert.AreEqual(true, shape.AspectRatioLocked);            
        }
#elif NETCOREAPP2_1 || __MOBILE__
        [Test]
        public void AspectRatioLockedDefaultValueNetStandard2()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // The best place for the watermark image is in the header or footer so it is shown on every page
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            
            using (SKManagedStream stream = new SKManagedStream(File.OpenRead(ImageDir + "Transparent background logo.png")))
            {
                using (SKBitmap bitmap = SKBitmap.Decode(stream))
                {
                    // Insert a floating picture.
                    Shape shape = builder.InsertImage(bitmap);
                    shape.WrapType = WrapType.None;
                    shape.BehindText = true;

                    shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
                    shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;

                    // Calculate image left and top position so it appears in the center of the page
                    shape.Left = (builder.PageSetup.PageWidth - shape.Width) / 2;
                    shape.Top = (builder.PageSetup.PageHeight - shape.Height) / 2;

                    doc = DocumentHelper.SaveOpen(doc);
        
                    shape = (Shape) doc.GetChild(NodeType.Shape, 0, true);
                    Assert.AreEqual(true, shape.AspectRatioLocked);
                }
            }            
        }
#endif

        [Test]
        public void Coordinates()
        {
            //ExStart
            //ExFor:ShapeBase.DistanceBottom
            //ExFor:ShapeBase.DistanceLeft
            //ExFor:ShapeBase.DistanceRight
            //ExFor:ShapeBase.DistanceTop
            //ExSummary:Shows how to set the wrapping distance for a text that surrounds a shape.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a rectangle and, get the text to wrap tightly around its bounds.
            Shape shape = builder.InsertShape(ShapeType.Rectangle, 150, 150);
            shape.WrapType = WrapType.Tight;

            // Set the minimum distance between the shape and surrounding text to 40pt from all sides.
            shape.DistanceTop = 40;
            shape.DistanceBottom = 40;
            shape.DistanceLeft = 40;
            shape.DistanceRight = 40;

            // Move the shape closer to the center of the page, and then rotate the shape 60 degrees clockwise.
            shape.Top = 75;
            shape.Left = 150; 
            shape.Rotation = 60;

            // Add text that will wrap around the shape.
            builder.Font.Size = 24;
            builder.Write("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
                          "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.");

            doc.Save(ArtifactsDir + "Shape.Coordinates.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.Coordinates.docx");
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyShape(ShapeType.Rectangle, "Rectangle 100002", 150.0d, 150.0d, 75.0d, 150.0d, shape);
            Assert.AreEqual(40.0d, shape.DistanceBottom);
            Assert.AreEqual(40.0d, shape.DistanceLeft);
            Assert.AreEqual(40.0d, shape.DistanceRight);
            Assert.AreEqual(40.0d, shape.DistanceTop);
            Assert.AreEqual(60.0d, shape.Rotation);
        }

        [Test]
        public void GroupShape()
        {
            //ExStart
            //ExFor:ShapeBase.Bounds
            //ExFor:ShapeBase.CoordOrigin
            //ExFor:ShapeBase.CoordSize
            //ExSummary:Shows how to create and populate a group shape.
            Document doc = new Document();

            // Create a group shape. A group shape can display a collection of child shape nodes.
            // In Microsoft Word, clicking within the group shape's boundary or on one of the group shape's child shapes will
            // select all the other child shapes within this group and allow us to scale and move all the shapes at once.
            GroupShape group = new GroupShape(doc);

            Assert.AreEqual(WrapType.None, group.WrapType);

            // Create a 400pt x 400pt group shape, and place it at the document's floating shape coordinate origin.
            group.Bounds = new RectangleF(0, 0, 400, 400);

            // Set the group's internal coordinate plane size to 500 x 500pt. 
            // The top left corner of the group will have an x and y coordinate of (0, 0),
            // and the bottom right corner will have an x and y coordinate of (500, 500).
            group.CoordSize = new Size(500, 500);

            // Set the coordinates of the top left corner of the group to (-250, -250). 
            // The group's center will now have an x and y coordinate value of (0, 0),
            // and the bottom right corner will be at (250, 250).
            group.CoordOrigin = new Point(-250, -250);

            // Create a rectangle that will display the boundary of this group shape, and add it to the group.
            group.AppendChild(new Shape(doc, ShapeType.Rectangle)
            {
                Width = group.CoordSize.Width,
                Height = group.CoordSize.Height,
                Left = group.CoordOrigin.X,
                Top = group.CoordOrigin.Y
            });

            // Once a shape is a part of a group shape, we can access it as a child node and then modify it.
            ((Shape)group.GetChild(NodeType.Shape, 0, true)).Stroke.DashStyle = DashStyle.Dash;

            // Create a small red star, and insert it into the group.
            // Line up the shape with the group's coordinate origin, which we have moved to the center.
            group.AppendChild(new Shape(doc, ShapeType.Star)
            {
                Width = 20,
                Height = 20,
                Left = -10,
                Top = -10,
                FillColor = Color.Red
            });

            // Insert a rectangle, and then insert a slightly smaller rectangle in the same place with an image. 
            // Newer shapes that we add to the group overlap older shapes. The light blue rectangle will partially overlap the red star,
            // and then the shape with the image will overlap the light blue rectangle, using it as a frame.
            // We cannot use the "ZOrder" properties of shapes to manipulate their arrangement within a group shape. 
            group.AppendChild(new Shape(doc, ShapeType.Rectangle)
            {
                Width = 250,
                Height = 250,
                Left = -250,
                Top = -250,
                FillColor = Color.LightBlue
            });

            group.AppendChild(new Shape(doc, ShapeType.Image)
            {
                Width = 200,
                Height = 200,
                Left = -225,
                Top = -225
            });

            ((Shape)group.GetChild(NodeType.Shape, 3, true)).ImageData.SetImage(ImageDir + "Logo.jpg");

            // Insert a text box into the group shape. Set the "Left" property so that the text box's right edge
            // touches the right boundary of the group shape. Set the "Top" property so that the text box sits outside
            // the boundary of the group shape, with its top size lined up along the group shape's bottom margin.
            group.AppendChild(new Shape(doc, ShapeType.TextBox)
            {
                Width = 200,
                Height = 50,
                Left = group.CoordSize.Width + group.CoordOrigin.X - 200,
                Top = group.CoordSize.Height + group.CoordOrigin.Y
            });

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertNode(group);
            builder.MoveTo(((Shape)group.GetChild(NodeType.Shape, 4, true)).AppendChild(new Paragraph(doc)));
            builder.Write("Hello world!");

            doc.Save(ArtifactsDir + "Shape.GroupShape.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.GroupShape.docx");
            group = (GroupShape)doc.GetChild(NodeType.GroupShape, 0, true);

            Assert.AreEqual(new RectangleF(0, 0, 400, 400), group.Bounds);
            Assert.AreEqual(new Size(500, 500), group.CoordSize);
            Assert.AreEqual(new Point(-250, -250), group.CoordOrigin);

            TestUtil.VerifyShape(ShapeType.Rectangle, string.Empty, 500.0d, 500.0d, -250.0d, -250.0d, (Shape)group.GetChild(NodeType.Shape, 0, true));
            TestUtil.VerifyShape(ShapeType.Star, string.Empty, 20.0d, 20.0d, -10.0d, -10.0d, (Shape)group.GetChild(NodeType.Shape, 1, true));
            TestUtil.VerifyShape(ShapeType.Rectangle, string.Empty, 250.0d, 250.0d, -250.0d, -250.0d, (Shape)group.GetChild(NodeType.Shape, 2, true));
            TestUtil.VerifyShape(ShapeType.Image, string.Empty, 200.0d, 200.0d, -225.0d, -225.0d, (Shape)group.GetChild(NodeType.Shape, 3, true));
            TestUtil.VerifyShape(ShapeType.TextBox, string.Empty, 200.0d, 50.0d, 250.0d, 50.0d, (Shape)group.GetChild(NodeType.Shape, 4, true));
        }

        [Test]
        public void IsTopLevel()
        {
            //ExStart
            //ExFor:ShapeBase.IsTopLevel
            //ExSummary:Shows how to tell whether a shape is a part of a group shape.
            Document doc = new Document();

            Shape shape = new Shape(doc, ShapeType.Rectangle);
            shape.Width = 200;
            shape.Height = 200;
            shape.WrapType = WrapType.None;

            // A shape by default is not part of any group shape, and therefore has the "IsTopLevel" property set to "true".
            Assert.True(shape.IsTopLevel);

            GroupShape group = new GroupShape(doc);
            group.AppendChild(shape);

            // Once we assimilate a shape into a group shape, the "IsTopLevel" property changes to "false".
            Assert.False(shape.IsTopLevel);
            //ExEnd
        }

        [Test]
        public void LocalToParent()
        {
            //ExStart
            //ExFor:ShapeBase.CoordOrigin
            //ExFor:ShapeBase.CoordSize
            //ExFor:ShapeBase.LocalToParent(PointF)
            //ExSummary:Shows how to translate the x and y coordinate location on a shape's coordinate plane to a location on the parent shape's coordinate plane.
            Document doc = new Document();

            // Insert a group shape, and place it 100 points below and to the right of
            // the document's x and Y coordinate origin point.
            GroupShape group = new GroupShape(doc);
            group.Bounds = new RectangleF(100, 100, 500, 500);

            // Use the "LocalToParent" method to determine that (0, 0) on the group's internal x and y coordinates
            // lies on (100, 100) of its parent shape's coordinate system. The group shape's parent is the document itself.
            Assert.AreEqual(new PointF(100, 100), group.LocalToParent(new PointF(0, 0)));

            // By default, a shape's internal coordinate plane has the top left corner at (0, 0),
            // and the bottom right corner at (1000, 1000). Due to its size, our group shape covers an area of 500pt x 500pt
            // in the document's plane. This means that a movement of 1pt on the document's coordinate plane will translate
            // to a movement of 2pts on the group shape's coordinate plane.
            Assert.AreEqual(new PointF(150, 150), group.LocalToParent(new PointF(100, 100)));
            Assert.AreEqual(new PointF(200, 200), group.LocalToParent(new PointF(200, 200)));
            Assert.AreEqual(new PointF(250, 250), group.LocalToParent(new PointF(300, 300)));

            // Move the group shape's x and y axis origin from the top left corner to the center.
            // This will offset the group's internal coordinates relative to the document's coordinates even further.
            group.CoordOrigin = new Point(-250, -250);

            Assert.AreEqual(new PointF(375, 375), group.LocalToParent(new PointF(300, 300)));

            // Changing the scale of the coordinate plane will also affect relative locations.
            group.CoordSize = new Size(500, 500);

            Assert.AreEqual(new PointF(650, 650), group.LocalToParent(new PointF(300, 300)));

            // If we wish to add a shape to this group while defining its location based on a location in the document,
            // we will need to first confirm a location in the group shape that will match the document's location.
            Assert.AreEqual(new PointF(700, 700), group.LocalToParent(new PointF(350, 350)));

            Shape shape = new Shape(doc, ShapeType.Rectangle)
            {
                Width = 100,
                Height = 100,
                Left = 700,
                Top = 700
            };

            group.AppendChild(shape);
            doc.FirstSection.Body.FirstParagraph.AppendChild(group);

            doc.Save(ArtifactsDir + "Shape.LocalToParent.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.LocalToParent.docx");
            group = (GroupShape)doc.GetChild(NodeType.GroupShape, 0, true);

            Assert.AreEqual(new RectangleF(100, 100, 500, 500), group.Bounds);
            Assert.AreEqual(new Size(500, 500), group.CoordSize);
            Assert.AreEqual(new Point(-250, -250), group.CoordOrigin);
        }

        [TestCase(false)]
        [TestCase(true)]
        public void AnchorLocked(bool anchorLocked)
        {
            //ExStart
            //ExFor:ShapeBase.AnchorLocked
            //ExSummary:Shows how to lock or unlock a shape's paragraph anchor.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello world!");

            builder.Write("Our shape will have an anchor attached to this paragraph.");
            Shape shape = builder.InsertShape(ShapeType.Rectangle, 200, 160);
            shape.WrapType = WrapType.None;
            builder.InsertBreak(BreakType.ParagraphBreak);

            builder.Writeln("Hello again!");

            // Set the "AnchorLocked" property to "true" to prevent the shape's anchor
            // from moving when moving the shape in Microsoft Word.
            // Set the "AnchorLocked" property to "false" to allow any movement of the shape
            // to also move its anchor to any other paragraph that the shape ends up close to.
            shape.AnchorLocked = anchorLocked;
            
            // If the shape does not have a visible anchor symbol to its left,
            // we will need to enable visible anchors via "Options" -> "Display" -> "Object Anchors".
            doc.Save(ArtifactsDir + "Shape.AnchorLocked.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.AnchorLocked.docx");
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            Assert.AreEqual(anchorLocked, shape.AnchorLocked);
        }

        [Test]
        public void DeleteAllShapes()
        {
            //ExStart
            //ExFor:Shape
            //ExSummary:Shows how to delete all shapes from a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert two shapes, and also a group shape with another shape inside it.
            builder.InsertShape(ShapeType.Rectangle, 400, 200);
            builder.InsertShape(ShapeType.Star, 300, 300);

            GroupShape group = new GroupShape(doc);
            group.Bounds = new RectangleF(100, 50, 200, 100);
            group.CoordOrigin = new Point(-1000, -500);

            Shape subShape = new Shape(doc, ShapeType.Cube);
            subShape.Width = 500;
            subShape.Height = 700;
            subShape.Left = 0;
            subShape.Top = 0;

            group.AppendChild(subShape);
            builder.InsertNode(group);

            Assert.AreEqual(3, doc.GetChildNodes(NodeType.Shape, true).Count);
            Assert.AreEqual(1, doc.GetChildNodes(NodeType.GroupShape, true).Count);

            // Remove all Shape nodes from the document.
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            shapes.Clear();

            // All shapes are gone, but the group shape is still in the document.
            Assert.AreEqual(1, doc.GetChildNodes(NodeType.GroupShape, true).Count);
            Assert.AreEqual(0, doc.GetChildNodes(NodeType.Shape, true).Count);

            // Remove all group shapes separately.
            NodeCollection groupShapes = doc.GetChildNodes(NodeType.GroupShape, true);
            groupShapes.Clear();

            Assert.AreEqual(0, doc.GetChildNodes(NodeType.GroupShape, true).Count);
            Assert.AreEqual(0, doc.GetChildNodes(NodeType.Shape, true).Count);
            //ExEnd
        }

        [Test]
        public void IsInline()
        {
            //ExStart
            //ExFor:ShapeBase.IsInline
            //ExSummary:Shows how to determine whether a shape is inline or floating.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Below are two wrapping types that shapes may have.
            // 1 -  Inline:
            builder.Write("Hello world! ");
            Shape shape = builder.InsertShape(ShapeType.Rectangle, 100, 100);
            shape.FillColor = Color.LightBlue;
            builder.Write(" Hello again.");

            // An inline shape sits inside a paragraph among other paragraph elements, such as runs of text.
            // In Microsoft Word, we may click and drag the shape to any paragraph as if it is a character.
            // If the shape is large, it will affect vertical paragraph spacing.
            // We cannot move this shape to a place with no paragraph.
            Assert.AreEqual(WrapType.Inline, shape.WrapType);
            Assert.True(shape.IsInline);

            // 2 -  Floating:
            shape = builder.InsertShape(ShapeType.Rectangle, RelativeHorizontalPosition.LeftMargin ,200, 
                RelativeVerticalPosition.TopMargin ,200, 100, 100, WrapType.None);
            shape.FillColor = Color.Orange;

            // A floating shape belongs to the paragraph that we insert it into,
            // which we can determine by an anchor symbol that appears when we click the shape.
            // If the shape does not have a visible anchor symbol to its left,
            // we will need to enable visible anchors via "Options" -> "Display" -> "Object Anchors".
            // In Microsoft Word, we may left click and drag this shape freely to any location.
            Assert.AreEqual(WrapType.None, shape.WrapType);
            Assert.False(shape.IsInline);

            doc.Save(ArtifactsDir + "Shape.IsInline.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.IsInline.docx");
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyShape(ShapeType.Rectangle, "Rectangle 100002", 100, 100, 0, 0, shape);
            Assert.AreEqual(Color.LightBlue.ToArgb(), shape.FillColor.ToArgb());
            Assert.AreEqual(WrapType.Inline, shape.WrapType);
            Assert.True(shape.IsInline);

            shape = (Shape)doc.GetChild(NodeType.Shape, 1, true);

            TestUtil.VerifyShape(ShapeType.Rectangle, "Rectangle 100004", 100, 100, 200, 200, shape);
            Assert.AreEqual(Color.Orange.ToArgb(), shape.FillColor.ToArgb());
            Assert.AreEqual(WrapType.None, shape.WrapType);
            Assert.False(shape.IsInline);
        }

        [Test]
        public void Bounds()
        {
            //ExStart
            //ExFor:ShapeBase.Bounds
            //ExFor:ShapeBase.BoundsInPoints
            //ExSummary:Shows how to verify shape containing block boundaries.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertShape(ShapeType.Line, RelativeHorizontalPosition.LeftMargin, 50,
                RelativeVerticalPosition.TopMargin, 50, 100, 100, WrapType.None);
            shape.StrokeColor = Color.Orange;

            // Even though the line itself takes up very little space on the document page,
            // it occupies a rectangular containing block, the size of which we can determine using the "Bounds" properties.
            Assert.AreEqual(new RectangleF(50, 50, 100, 100), shape.Bounds);
            Assert.AreEqual(new RectangleF(50, 50, 100, 100), shape.BoundsInPoints);

            // Create a group shape, and then set the size of its containing block using the "Bounds" property.
            GroupShape group = new GroupShape(doc);
            group.Bounds = new RectangleF(0, 100, 250, 250);

            Assert.AreEqual(new RectangleF(0, 100, 250, 250), group.BoundsInPoints);

            // Create a rectangle, verify the size of its bounding block, and then add it to the group shape.
            shape = new Shape(doc, ShapeType.Rectangle)
            {
                Width = 100,
                Height = 100,
                Left = 700,
                Top = 700
            };

            Assert.AreEqual(new RectangleF(700, 700, 100, 100), shape.BoundsInPoints);

            group.AppendChild(shape);

            // The group shape's coordinate plane has its origin on the top left-hand side corner of its containing block,
            // and the x and y coordinates of (1000, 1000) on the bottom right-hand side corner.
            // Our group shape is 250x250pt in size, so every 4pt on the group shape's coordinate plane
            // translates to 1pt in the document body's coordinate plane.
            // Every shape that we insert will also shrink in size by a factor of 4.
            // The change in the shape's "BoundsInPoints" property will reflect this.
            Assert.AreEqual(new RectangleF(175, 275, 25, 25), shape.BoundsInPoints);

            doc.FirstSection.Body.FirstParagraph.AppendChild(group);

            // Insert a shape, and place it outside of the bounds of the group shape's containing block.
            shape = new Shape(doc, ShapeType.Rectangle)
            {
                Width = 100,
                Height = 100,
                Left = 1000,
                Top = 1000
            };

            group.AppendChild(shape);

            // The group shape's footprint in the document body has increased, but the containing block remains the same.
            Assert.AreEqual(new RectangleF(0, 100, 250, 250), group.BoundsInPoints);
            Assert.AreEqual(new RectangleF(250, 350, 25, 25), shape.BoundsInPoints);

            doc.Save(ArtifactsDir + "Shape.Bounds.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.Bounds.docx");
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyShape(ShapeType.Line, "Line 100002", 100, 100, 50, 50, shape);
            Assert.AreEqual(Color.Orange.ToArgb(), shape.StrokeColor.ToArgb());
            Assert.AreEqual(new RectangleF(50, 50, 100, 100), shape.BoundsInPoints);

            group = (GroupShape)doc.GetChild(NodeType.GroupShape, 0, true);

            Assert.AreEqual(new RectangleF(0, 100, 250, 250), group.Bounds);
            Assert.AreEqual(new RectangleF(0, 100, 250, 250), group.BoundsInPoints);
            Assert.AreEqual(new Size(1000, 1000), group.CoordSize);
            Assert.AreEqual(new Point(0, 0), group.CoordOrigin);

            shape = (Shape)doc.GetChild(NodeType.Shape, 1, true);

            TestUtil.VerifyShape(ShapeType.Rectangle, string.Empty, 100, 100, 700, 700, shape);
            Assert.AreEqual(new RectangleF(175, 275, 25, 25), shape.BoundsInPoints);

            shape = (Shape)doc.GetChild(NodeType.Shape, 2, true);

            TestUtil.VerifyShape(ShapeType.Rectangle, string.Empty, 100, 100, 1000, 1000, shape);
            Assert.AreEqual(new RectangleF(250, 350, 25, 25), shape.BoundsInPoints);
        }

        [Test]
        public void FlipShapeOrientation()
        {
            //ExStart
            //ExFor:ShapeBase.FlipOrientation
            //ExFor:FlipOrientation
            //ExSummary:Shows how to flip a shape on an axis.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert an image shape, and leave its orientation in its default state.
            Shape shape = builder.InsertShape(ShapeType.Rectangle, RelativeHorizontalPosition.LeftMargin, 100,
                RelativeVerticalPosition.TopMargin, 100, 100, 100, WrapType.None);
            shape.ImageData.SetImage(ImageDir + "Logo.jpg");

            Assert.AreEqual(FlipOrientation.None, shape.FlipOrientation);

            shape = builder.InsertShape(ShapeType.Rectangle, RelativeHorizontalPosition.LeftMargin, 250,
                RelativeVerticalPosition.TopMargin, 100, 100, 100, WrapType.None);
            shape.ImageData.SetImage(ImageDir + "Logo.jpg");

            // Set the "FlipOrientation" property to "FlipOrientation.Horizontal" to flip the second shape on the y-axis,
            // making it into a horizontal mirror image of the first shape.
            shape.FlipOrientation = FlipOrientation.Horizontal;

            shape = builder.InsertShape(ShapeType.Rectangle, RelativeHorizontalPosition.LeftMargin, 100,
                RelativeVerticalPosition.TopMargin, 250, 100, 100, WrapType.None);
            shape.ImageData.SetImage(ImageDir + "Logo.jpg");

            // Set the "FlipOrientation" property to "FlipOrientation.Horizontal" to flip the third shape on the x-axis,
            // making it into a vertical mirror image of the first shape.
            shape.FlipOrientation = FlipOrientation.Vertical;

            shape = builder.InsertShape(ShapeType.Rectangle, RelativeHorizontalPosition.LeftMargin, 250,
                RelativeVerticalPosition.TopMargin, 250, 100, 100, WrapType.None);
            shape.ImageData.SetImage(ImageDir + "Logo.jpg");

            // Set the "FlipOrientation" property to "FlipOrientation.Horizontal" to flip the fourth shape on both the x and y axes,
            // making it into a horizontal and vertical mirror image of the first shape.
            shape.FlipOrientation = FlipOrientation.Both;
            
            doc.Save(ArtifactsDir + "Shape.FlipShapeOrientation.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.FlipShapeOrientation.docx");
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyShape(ShapeType.Rectangle, "Rectangle 100002", 100, 100, 100, 100, shape);
            Assert.AreEqual(FlipOrientation.None, shape.FlipOrientation);

            shape = (Shape)doc.GetChild(NodeType.Shape, 1, true);

            TestUtil.VerifyShape(ShapeType.Rectangle, "Rectangle 100004", 100, 100, 100, 250, shape);
            Assert.AreEqual(FlipOrientation.Horizontal, shape.FlipOrientation);

            shape = (Shape)doc.GetChild(NodeType.Shape, 2, true);

            TestUtil.VerifyShape(ShapeType.Rectangle, "Rectangle 100006", 100, 100, 250, 100, shape);
            Assert.AreEqual(FlipOrientation.Vertical, shape.FlipOrientation);

            shape = (Shape)doc.GetChild(NodeType.Shape, 3, true);

            TestUtil.VerifyShape(ShapeType.Rectangle, "Rectangle 100008", 100, 100, 250, 250, shape);
            Assert.AreEqual(FlipOrientation.Both, shape.FlipOrientation);
        }

        [Test]
        public void Fill()
        {
            //ExStart
            //ExFor:Shape.Fill
            //ExFor:Shape.FillColor
            //ExFor:Shape.StrokeColor
            //ExFor:Fill
            //ExFor:Fill.Opacity
            //ExSummary:Shows how to fill a shape with a solid color.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write some text, and then cover it with a floating shape.
            builder.Font.Size = 32;
            builder.Writeln("Hello world!");

            Shape shape = builder.InsertShape(ShapeType.CloudCallout, RelativeHorizontalPosition.LeftMargin, 25,
                RelativeVerticalPosition.TopMargin, 25, 250, 150, WrapType.None);

            // Use the "StrokeColor" property to set the color of the outline of the shape.
            shape.StrokeColor = Color.CadetBlue;

            // Use the "FillColor" property to set the color of the inside area of the shape.
            shape.FillColor = Color.LightBlue;

            // The "Opacity" property determines how transparent the color is on a 0-1 scale,
            // with 1 being fully opaque, and 0 being invisible.
            // The shape fill by default is fully opaque, so we cannot see the text that this shape is on top of.
            Assert.AreEqual(1.0d, shape.Fill.Opacity);

            // Set the shape fill color's opacity to a lower value so that we can see the text underneath it.
            shape.Fill.Opacity = 0.3;

            doc.Save(ArtifactsDir + "Shape.Fill.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.Fill.docx");
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyShape(ShapeType.CloudCallout, "CloudCallout 100002", 250.0d, 150.0d, 25.0d, 25.0d, shape);
            Assert.AreEqual(Color.LightBlue.ToArgb(), shape.FillColor.ToArgb());
            Assert.AreEqual(Color.CadetBlue.ToArgb(), shape.StrokeColor.ToArgb());
            Assert.AreEqual(0.3d, shape.Fill.Opacity, 0.01d);
        }

        [Test]
        public void Title()
        {
            //ExStart
            //ExFor:ShapeBase.Title
            //ExSummary:Shows how to set the title of a shape.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a shape, give it a title, and then add it to the document.
            Shape shape = new Shape(doc, ShapeType.Cube);
            shape.Width = 200;
            shape.Height = 200;
            shape.Title = "My cube";
            
            builder.InsertNode(shape);

            // When we save a document with a shape that has a title,
            // Aspose.Words will store that title in the shape's Alt Text.
            doc.Save(ArtifactsDir + "Shape.Title.docx");

            doc = new Document(ArtifactsDir + "Shape.Title.docx");
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            Assert.AreEqual(string.Empty, shape.Title);
            Assert.AreEqual("Title: My cube", shape.AlternativeText);
            //ExEnd

            TestUtil.VerifyShape(ShapeType.Cube, string.Empty, 200.0d, 200.0d, 0.0d, 0.0d, shape);
        }

        [Test]
        public void ReplaceTextboxesWithImages()
        {
            //ExStart
            //ExFor:WrapSide
            //ExFor:ShapeBase.WrapSide
            //ExFor:NodeCollection
            //ExFor:CompositeNode.InsertAfter(Node, Node)
            //ExFor:NodeCollection.ToArray
            //ExSummary:Shows how to replace all textbox shapes with image shapes.
            Document doc = new Document(MyDir + "Textboxes in drawing canvas.docx");

            Shape[] shapes = doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>().ToArray();

            Assert.AreEqual(3, shapes.Count(s => s.ShapeType == ShapeType.TextBox));
            Assert.AreEqual(1, shapes.Count(s => s.ShapeType == ShapeType.Image));

            foreach (Shape shape in shapes)
            {
                if (shape.ShapeType.Equals(ShapeType.TextBox))
                {
                    Shape replacementShape = new Shape(doc, ShapeType.Image);
                    replacementShape.ImageData.SetImage(ImageDir + "Logo.jpg");
                    replacementShape.Left = shape.Left;
                    replacementShape.Top = shape.Top;
                    replacementShape.Width = shape.Width;
                    replacementShape.Height = shape.Height;
                    replacementShape.RelativeHorizontalPosition = shape.RelativeHorizontalPosition;
                    replacementShape.RelativeVerticalPosition = shape.RelativeVerticalPosition;
                    replacementShape.HorizontalAlignment = shape.HorizontalAlignment;
                    replacementShape.VerticalAlignment = shape.VerticalAlignment;
                    replacementShape.WrapType = shape.WrapType;
                    replacementShape.WrapSide = shape.WrapSide;

                    shape.ParentNode.InsertAfter(replacementShape, shape);
                    shape.Remove();
                }
            }

            shapes = doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>().ToArray();

            Assert.AreEqual(0, shapes.Count(s => s.ShapeType == ShapeType.TextBox));
            Assert.AreEqual(4, shapes.Count(s => s.ShapeType == ShapeType.Image));

            doc.Save(ArtifactsDir + "Shape.ReplaceTextboxesWithImages.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.ReplaceTextboxesWithImages.docx");
            Shape outShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            Assert.AreEqual(WrapSide.Both, outShape.WrapSide);
        }

        [Test]
        public void CreateTextBox()
        {
            //ExStart
            //ExFor:Shape.#ctor(DocumentBase, ShapeType)
            //ExFor:Story.FirstParagraph
            //ExFor:Shape.FirstParagraph
            //ExFor:ShapeBase.WrapType
            //ExSummary:Shows how to create and format a text box.
            Document doc = new Document();

            // Create a floating text box.
            Shape textBox = new Shape(doc, ShapeType.TextBox);
            textBox.WrapType = WrapType.None;
            textBox.Height = 50;
            textBox.Width = 200;

            // Set the horizontal, and vertical alignment of the text inside the shape.
            textBox.HorizontalAlignment = HorizontalAlignment.Center;
            textBox.VerticalAlignment = VerticalAlignment.Top;
            
            // Add a paragraph to the text box, and add a run of text that the text box will display.
            textBox.AppendChild(new Paragraph(doc));
            Paragraph para = textBox.FirstParagraph;
            para.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            Run run = new Run(doc);
            run.Text = "Hello world!";
            para.AppendChild(run);

            doc.FirstSection.Body.FirstParagraph.AppendChild(textBox);

            doc.Save(ArtifactsDir + "Shape.CreateTextBox.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.CreateTextBox.docx");
            textBox = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyShape(ShapeType.TextBox, string.Empty, 200.0d, 50.0d, 0.0d, 0.0d, textBox);
            Assert.AreEqual(WrapType.None, textBox.WrapType);
            Assert.AreEqual(HorizontalAlignment.Center, textBox.HorizontalAlignment);
            Assert.AreEqual(VerticalAlignment.Top, textBox.VerticalAlignment);
            Assert.AreEqual("Hello world!", textBox.GetText().Trim());
        }

        [Test]
        public void ZOrder()
        {
            //ExStart
            //ExFor:ShapeBase.ZOrder
            //ExSummary:Shows how to manipulate the order of shapes.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert three different colored rectangles that partially overlap each other.
            // When we insert a shape that overlaps another shape, Aspose.Words places the newer shape on top of the old one.
            // The light green rectangle will overlap the light blue rectangle and partially obscure it,
            // and the light blue rectangle will obscure the orange rectangle.
            Shape shape = builder.InsertShape(ShapeType.Rectangle, RelativeHorizontalPosition.LeftMargin, 100,
                RelativeVerticalPosition.TopMargin, 100, 200, 200, WrapType.None);
            shape.FillColor = Color.Orange;

            shape = builder.InsertShape(ShapeType.Rectangle, RelativeHorizontalPosition.LeftMargin, 150,
                RelativeVerticalPosition.TopMargin, 150, 200, 200, WrapType.None);
            shape.FillColor = Color.LightBlue;

            shape = builder.InsertShape(ShapeType.Rectangle, RelativeHorizontalPosition.LeftMargin, 200,
                RelativeVerticalPosition.TopMargin, 200, 200, 200, WrapType.None);
            shape.FillColor = Color.LightGreen;

            Shape[] shapes = doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>().ToArray();

            // The "ZOrder" property of a shape determines its stacking priority among other overlapping shapes.
            // If two overlapping shapes have different "ZOrder" values,
            // Microsoft Word will place the shape with the higher value over the shape with the lower value. 
            // Set the "ZOrder" values of our shapes to place the first orange rectangle over the second light blue one,
            // and the second light blue rectangle over the third light green rectangle.
            // This will reverse their original stacking order.
            shapes[0].ZOrder = 3;
            shapes[1].ZOrder = 2;
            shapes[2].ZOrder = 1;

            doc.Save(ArtifactsDir + "Shape.ZOrder.docx");
            //ExEnd
        }

        [Test]
        public void GetActiveXControlProperties()
        {
            //ExStart
            //ExFor:OleControl
            //ExFor:Ole.OleControl.IsForms2OleControl
            //ExFor:Ole.OleControl.Name
            //ExFor:OleFormat.OleControl
            //ExFor:Forms2OleControl
            //ExFor:Forms2OleControl.Caption
            //ExFor:Forms2OleControl.Value
            //ExFor:Forms2OleControl.Enabled
            //ExFor:Forms2OleControl.Type
            //ExFor:Forms2OleControl.ChildNodes
            //ExSummary:Shows how to verify the properties of an ActiveX control.
            Document doc = new Document(MyDir + "ActiveX controls.docx");

            Shape shape = (Shape) doc.GetChild(NodeType.Shape, 0, true);
            OleControl oleControl = shape.OleFormat.OleControl;

            Assert.AreEqual(null, oleControl.Name);

            if (oleControl.IsForms2OleControl)
            {
                Forms2OleControl checkBox = (Forms2OleControl) oleControl;
                Assert.AreEqual("Первый", checkBox.Caption);
                Assert.AreEqual("0", checkBox.Value);
                Assert.AreEqual(true, checkBox.Enabled);
                Assert.AreEqual(Forms2OleControlType.CheckBox, checkBox.Type);
                Assert.AreEqual(null, checkBox.ChildNodes);
            }
            //ExEnd
        }

        [Test]
        public void GetOleObjectRawData()
        {
            //ExStart
            //ExFor:OleFormat.GetRawData
            //ExSummary:Shows how to access the raw data of an embedded OLE object.
            Document doc = new Document(MyDir + "OLE objects.docx");

            foreach (Node shape in doc.GetChildNodes(NodeType.Shape, true))
            {
                OleFormat oleFormat = ((Shape)shape).OleFormat;
                if (oleFormat != null)
                {
                    Console.WriteLine($"This is {(oleFormat.IsLink ? "a linked" : "an embedded")} object");
                    byte[] oleRawData = oleFormat.GetRawData();

                    Assert.AreEqual(24576, oleRawData.Length);
                }
            }
            //ExEnd
        }

        [Test]
        public void OleControl()
        {
            //ExStart
            //ExFor:OleFormat
            //ExFor:OleFormat.AutoUpdate
            //ExFor:OleFormat.IsLocked
            //ExFor:OleFormat.ProgId
            //ExFor:OleFormat.Save(Stream)
            //ExFor:OleFormat.Save(String)
            //ExFor:OleFormat.SuggestedExtension
            //ExSummary:Shows how to extract embedded OLE objects into files.
            Document doc = new Document(MyDir + "OLE spreadsheet.docm");
            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            // The OLE object in the first shape is a Microsoft Excel spreadsheet.
            OleFormat oleFormat = shape.OleFormat;
            Assert.AreEqual("Excel.Sheet.12", oleFormat.ProgId);

            // Our object is neither auto updating nor locked from updates.
            Assert.False(oleFormat.AutoUpdate);
            Assert.AreEqual(false, oleFormat.IsLocked);

            // If we plan on saving the OLE object to a file in the local file system,
            // we can use the "SuggestedExtension" property to determine which file extension to apply to the file.
            Assert.AreEqual(".xlsx", oleFormat.SuggestedExtension);

            // Below are two ways of saving an OLE object to a file in the local file system.
            // 1 -  Save it via a stream:
            using (FileStream fs = new FileStream(ArtifactsDir + "OLE spreadsheet extracted via stream" + oleFormat.SuggestedExtension, FileMode.Create))
            {
                oleFormat.Save(fs);
            }

            // 2 -  Save it directly to a filename:
            oleFormat.Save(ArtifactsDir + "OLE spreadsheet saved directly" + oleFormat.SuggestedExtension);
            //ExEnd

            Assert.That(8000, Is.LessThan(new FileInfo(ArtifactsDir + "OLE spreadsheet extracted via stream.xlsx").Length));
            Assert.That(8000, Is.LessThan(new FileInfo(ArtifactsDir + "OLE spreadsheet saved directly.xlsx").Length));
        }

        [Test]
        public void OleLinks()
        {
            //ExStart
            //ExFor:OleFormat.IconCaption
            //ExFor:OleFormat.GetOleEntry(String)
            //ExFor:OleFormat.IsLink
            //ExFor:OleFormat.OleIcon
            //ExFor:OleFormat.SourceFullName
            //ExFor:OleFormat.SourceItem
            //ExSummary:Shows how to insert linked and unlinked OLE objects.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Embed a Microsoft Visio drawing as an OLE object into the document
            builder.InsertOleObject(ImageDir + "Microsoft Visio drawing.vsd", "Package", false, false, null);

            // Insert a link to the file in the local file system and display it as an icon
            builder.InsertOleObject(ImageDir + "Microsoft Visio drawing.vsd", "Package", true, true, null);
            
            // Both the OLE objects are stored within shapes
            List<Shape> shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().ToList();
            Assert.AreEqual(2, shapes.Count);

            // If the shape is an OLE object, it will have a valid OleFormat property
            // We can use it check if it is linked or displayed as an icon, among other things
            OleFormat oleFormat = shapes[0].OleFormat;
            Assert.AreEqual(false, oleFormat.IsLink);
            Assert.AreEqual(false, oleFormat.OleIcon);

            oleFormat = shapes[1].OleFormat;
            Assert.AreEqual(true, oleFormat.IsLink);
            Assert.AreEqual(true, oleFormat.OleIcon);

            // Get the name or the source file and verify that the whole file is linked
            Assert.True(oleFormat.SourceFullName.EndsWith(@"Images" + Path.DirectorySeparatorChar + "Microsoft Visio drawing.vsd"));
            Assert.AreEqual("", oleFormat.SourceItem);

            Assert.AreEqual("Packager", oleFormat.IconCaption);

            doc.Save(ArtifactsDir + "Shape.OleLinks.docx");

            // If the object has OLE data, we can access it in the form of a stream
            using (MemoryStream stream = oleFormat.GetOleEntry("\x0001CompObj"))
            {
                byte[] oleEntryBytes = stream.ToArray();
                Assert.AreEqual(76, oleEntryBytes.Length);
            }
            //ExEnd
        }

        [Test]
        public void OleControlCollection()
        {
            //ExStart
            //ExFor:OleFormat.Clsid
            //ExFor:Ole.Forms2OleControlCollection
            //ExFor:Ole.Forms2OleControlCollection.Count
            //ExFor:Ole.Forms2OleControlCollection.Item(Int32)
            //ExSummary:Shows how to access an OLE control embedded in a document and its child controls.
            // Open a document that contains a Microsoft Forms OLE control with child controls
            Document doc = new Document(MyDir + "OLE ActiveX controls.docm");

            // Get the shape that contains the control
            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            Assert.AreEqual("6e182020-f460-11ce-9bcd-00aa00608e01", shape.OleFormat.Clsid.ToString());

            Forms2OleControl oleControl = (Forms2OleControl)shape.OleFormat.OleControl;

            // Some controls contain child controls
            Forms2OleControlCollection oleControlCollection = oleControl.ChildNodes;

            // In this case, the child controls are 3 option buttons
            Assert.AreEqual(3, oleControlCollection.Count);

            Assert.AreEqual("C#", oleControlCollection[0].Caption);
            Assert.AreEqual("1", oleControlCollection[0].Value);

            Assert.AreEqual("Visual Basic", oleControlCollection[1].Caption);
            Assert.AreEqual("0", oleControlCollection[1].Value);

            Assert.AreEqual("Delphi", oleControlCollection[2].Caption);
            Assert.AreEqual("0", oleControlCollection[2].Value);
            //ExEnd
        }

        [Test]
        public void SuggestedFileName()
        {
            //ExStart
            //ExFor:OleFormat.SuggestedFileName
            //ExSummary:Shows how to get suggested file name from the object.
            Document doc = new Document(MyDir + "OLE shape.rtf");

            // Gets the file name suggested for the current embedded object if you want to save it into a file
            Shape oleShape = (Shape) doc.FirstSection.Body.GetChild(NodeType.Shape, 0, true);
            string suggestedFileName = oleShape.OleFormat.SuggestedFileName;

            Assert.AreEqual("CSV.csv", suggestedFileName);
            //ExEnd
        }

        [Test]
        public void ObjectDidNotHaveSuggestedFileName()
        {
            Document doc = new Document(MyDir + "ActiveX controls.docx");

            Shape shape = (Shape) doc.GetChild(NodeType.Shape, 0, true);
            Assert.That(shape.OleFormat.SuggestedFileName, Is.Empty);
        }

        [Test]
        public void ResolutionDefaultValues()
        {
            ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.Jpeg);

            Assert.AreEqual(96, imageOptions.HorizontalResolution);
            Assert.AreEqual(96, imageOptions.VerticalResolution);
        }

        [Test]
        public void SaveShapeObjectAsImage()
        {
            //ExStart
            //ExFor:OfficeMath.GetMathRenderer
            //ExFor:NodeRendererBase.Save(String, ImageSaveOptions)
            //ExSummary:Shows how to convert specific object into image
            Document doc = new Document(MyDir + "Office math.docx");

            // Get OfficeMath node from the document and render this as image (you can also do the same with the Shape node)
            OfficeMath math = (OfficeMath)doc.GetChild(NodeType.OfficeMath, 0, true);
            math.GetMathRenderer().Save(ArtifactsDir + "Shape.SaveShapeObjectAsImage.png", new ImageSaveOptions(SaveFormat.Png));
            //ExEnd

            if (!IsRunningOnMono())
                TestUtil.VerifyImage(159, 18, ArtifactsDir + "Shape.SaveShapeObjectAsImage.png");
            else
                TestUtil.VerifyImage(147, 26, ArtifactsDir + "Shape.SaveShapeObjectAsImage.png");
        }

        [Test]
        public void OfficeMathDisplayException()
        {
            Document doc = new Document(MyDir + "Office math.docx");

            OfficeMath officeMath = (OfficeMath) doc.GetChild(NodeType.OfficeMath, 0, true);
            officeMath.DisplayType = OfficeMathDisplayType.Display;

            Assert.That(() => officeMath.Justification = OfficeMathJustification.Inline,
                Throws.TypeOf<ArgumentException>());
        }

        [Test]
        public void OfficeMathDefaultValue()
        {
            Document doc = new Document(MyDir + "Office math.docx");

            OfficeMath officeMath = (OfficeMath) doc.GetChild(NodeType.OfficeMath, 6, true);

            Assert.AreEqual(OfficeMathDisplayType.Inline, officeMath.DisplayType);
            Assert.AreEqual(OfficeMathJustification.Inline, officeMath.Justification);
        }

        [Test]
        public void OfficeMath()
        {
            //ExStart
            //ExFor:OfficeMath
            //ExFor:OfficeMath.DisplayType
            //ExFor:OfficeMath.EquationXmlEncoding
            //ExFor:OfficeMath.Justification
            //ExFor:OfficeMath.NodeType
            //ExFor:OfficeMath.ParentParagraph
            //ExFor:OfficeMathDisplayType
            //ExFor:OfficeMathJustification
            //ExSummary:Shows how to set office math display formatting.
            Document doc = new Document(MyDir + "Office math.docx");

            OfficeMath officeMath = (OfficeMath) doc.GetChild(NodeType.OfficeMath, 0, true);

            // OfficeMath nodes that are children of other OfficeMath nodes are always inline
            // The node we are working with is a base node, so its location and display type can be changed
            Assert.AreEqual(MathObjectType.OMathPara, officeMath.MathObjectType);
            Assert.AreEqual(NodeType.OfficeMath, officeMath.NodeType);
            Assert.AreEqual(officeMath.ParentNode, officeMath.ParentParagraph);

            // Used by OOXML and WML formats
            Assert.IsNull(officeMath.EquationXmlEncoding);

            // We can change the location and display type of the OfficeMath node
            officeMath.DisplayType = OfficeMathDisplayType.Display;
            officeMath.Justification = OfficeMathJustification.Left;

            doc.Save(ArtifactsDir + "Shape.OfficeMath.docx");
            //ExEnd

            Assert.IsTrue(DocumentHelper.CompareDocs(ArtifactsDir + "Shape.OfficeMath.docx", GoldsDir + "Shape.OfficeMath Gold.docx"));
        }

        [Test]
        public void CannotBeSetDisplayWithInlineJustification()
        {
            Document doc = new Document(MyDir + "Office math.docx");

            OfficeMath officeMath = (OfficeMath) doc.GetChild(NodeType.OfficeMath, 0, true);
            officeMath.DisplayType = OfficeMathDisplayType.Display;

            Assert.Throws<ArgumentException>(() => officeMath.Justification = OfficeMathJustification.Inline);
        }

        [Test]
        public void CannotBeSetInlineDisplayWithJustification()
        {
            Document doc = new Document(MyDir + "Office math.docx");

            OfficeMath officeMath = (OfficeMath) doc.GetChild(NodeType.OfficeMath, 0, true);
            officeMath.DisplayType = OfficeMathDisplayType.Inline;

            Assert.Throws<ArgumentException>(() => officeMath.Justification = OfficeMathJustification.Center);
        }

        [Test]
        public void OfficeMathDisplayNestedObjects()
        {
            Document doc = new Document(MyDir + "Office math.docx");

            OfficeMath officeMath = (OfficeMath) doc.GetChild(NodeType.OfficeMath, 0, true);

            // Always inline
            Assert.AreEqual(OfficeMathDisplayType.Display, officeMath.DisplayType);
            Assert.AreEqual(OfficeMathJustification.Center, officeMath.Justification);
        }

        [TestCase(0, MathObjectType.OMathPara)]
        [TestCase(1, MathObjectType.OMath)]
        [TestCase(2, MathObjectType.Supercript)]
        [TestCase(3, MathObjectType.Argument)]
        [TestCase(4, MathObjectType.SuperscriptPart)]
        public void WorkWithMathObjectType(int index, MathObjectType objectType)
        {
            Document doc = new Document(MyDir + "Office math.docx");

            OfficeMath officeMath = (OfficeMath) doc.GetChild(NodeType.OfficeMath, index, true);
            Assert.AreEqual(objectType, officeMath.MathObjectType);
        }

        [TestCase(true)]
        [TestCase(false)]
        public void AspectRatioLocked(bool isLocked)
        {
            //ExStart
            //ExFor:ShapeBase.AspectRatioLocked
            //ExSummary:Shows how to set "AspectRatioLocked" for the shape object.
            Document doc = new Document(MyDir + "ActiveX controls.docx");

            // Get shape object from the document and set AspectRatioLocked,
            // which is affects only top level shapes, to mimic Microsoft Word behavior
            Shape shape = (Shape) doc.GetChild(NodeType.Shape, 0, true);
            shape.AspectRatioLocked = isLocked;
            //ExEnd

            doc = DocumentHelper.SaveOpen(doc);
            shape = (Shape) doc.GetChild(NodeType.Shape, 0, true);

            Assert.AreEqual(isLocked, shape.AspectRatioLocked);
        }

        [Test]
        public void MarkupLunguageByDefault()
        {
            //ExStart
            //ExFor:ShapeBase.MarkupLanguage
            //ExFor:ShapeBase.SizeInPoints
            //ExSummary:Shows how get markup language for shape object in document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertImage(ImageDir + "Transparent background logo.png");

            // Loop through all single shapes inside document
            foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>())
            {
                Assert.AreEqual(ShapeMarkupLanguage.Dml, shape.MarkupLanguage); //ExSkip

                Console.WriteLine("Shape: " + shape.MarkupLanguage);
                Console.WriteLine("ShapeSize: " + shape.SizeInPoints);
            }
            //ExEnd
        }

        [TestCase(MsWordVersion.Word2000, ShapeMarkupLanguage.Vml)]
        [TestCase(MsWordVersion.Word2002, ShapeMarkupLanguage.Vml)]
        [TestCase(MsWordVersion.Word2003, ShapeMarkupLanguage.Vml)]
        [TestCase(MsWordVersion.Word2007, ShapeMarkupLanguage.Vml)]
        [TestCase(MsWordVersion.Word2010, ShapeMarkupLanguage.Dml)]
        [TestCase(MsWordVersion.Word2013, ShapeMarkupLanguage.Dml)]
        [TestCase(MsWordVersion.Word2016, ShapeMarkupLanguage.Dml)]
        public void MarkupLunguageForDifferentMsWordVersions(MsWordVersion msWordVersion,
            ShapeMarkupLanguage shapeMarkupLanguage)
        {
            Document doc = new Document();
            doc.CompatibilityOptions.OptimizeFor(msWordVersion);

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertImage(ImageDir + "Transparent background logo.png");

            // Loop through all single shapes inside document
            foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>())
            {
                Assert.AreEqual(shapeMarkupLanguage, shape.MarkupLanguage);
            }
        }

        [Test]
        public void ChangeStrokeProperties()
        {
            //ExStart
            //ExFor:Stroke
            //ExFor:Stroke.On
            //ExFor:Stroke.Weight
            //ExFor:Stroke.JoinStyle
            //ExFor:Stroke.LineStyle
            //ExFor:ShapeLineStyle
            //ExSummary:Shows how change stroke properties.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a new shape of type Rectangle
            Shape rectangle = new Shape(doc, ShapeType.Rectangle);

            // Change stroke properties
            Stroke stroke = rectangle.Stroke;
            stroke.On = true;
            stroke.Weight = 5;
            stroke.Color = Color.Red;
            stroke.DashStyle = DashStyle.ShortDashDotDot;
            stroke.JoinStyle = JoinStyle.Miter;
            stroke.EndCap = EndCap.Square;
            stroke.LineStyle = ShapeLineStyle.Triple;

            // Insert shape object
            builder.InsertNode(rectangle);
            //ExEnd

            doc = DocumentHelper.SaveOpen(doc);
            rectangle = (Shape) doc.GetChild(NodeType.Shape, 0, true);
            Stroke strokeAfter = rectangle.Stroke;

            Assert.AreEqual(true, strokeAfter.On);
            Assert.AreEqual(5, strokeAfter.Weight);
            Assert.AreEqual(Color.Red.ToArgb(), strokeAfter.Color.ToArgb());
            Assert.AreEqual(DashStyle.ShortDashDotDot, strokeAfter.DashStyle);
            Assert.AreEqual(JoinStyle.Miter, strokeAfter.JoinStyle);
            Assert.AreEqual(EndCap.Square, strokeAfter.EndCap);
            Assert.AreEqual(ShapeLineStyle.Triple, strokeAfter.LineStyle);
        }

        [Test, Description("WORDSNET-16067")]
        public void InsertOleObjectAsHtmlFile()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertOleObject("http://www.aspose.com", "htmlfile", true, false, null);

            doc.Save(ArtifactsDir + "Shape.InsertOleObjectAsHtmlFile.docx");
        }

        [Test, Description("WORDSNET-16085")]
        public void InsertOlePackage()
        {
            //ExStart
            //ExFor:OlePackage
            //ExFor:OleFormat.OlePackage
            //ExFor:OlePackage.FileName
            //ExFor:OlePackage.DisplayName
            //ExSummary:Shows how insert ole object as ole package and set it file name and display name.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            byte[] zipFileBytes = File.ReadAllBytes(DatabaseDir + "cat001.zip");

            using (MemoryStream stream = new MemoryStream(zipFileBytes))
            {
                Shape shape = builder.InsertOleObject(stream, "Package", true, null);

                OlePackage setOlePackage = shape.OleFormat.OlePackage;
                setOlePackage.FileName = "Cat FileName.zip";
                setOlePackage.DisplayName = "Cat DisplayName.zip";

                doc.Save(ArtifactsDir + "Shape.InsertOlePackage.docx");
            }
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.InsertOlePackage.docx");

            Shape getShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            OlePackage getOlePackage = getShape.OleFormat.OlePackage;

            Assert.AreEqual("Cat FileName.zip", getOlePackage.FileName);
            Assert.AreEqual("Cat DisplayName.zip", getOlePackage.DisplayName);
        }

        [Test]
        public void GetAccessToOlePackage()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape oleObject = builder.InsertOleObject(MyDir + "Spreadsheet.xlsx", false, false, null);
            Shape oleObjectAsOlePackage =
                builder.InsertOleObject(MyDir + "Spreadsheet.xlsx", "Excel.Sheet", false, false, null);

            Assert.AreEqual(null, oleObject.OleFormat.OlePackage);
            Assert.AreEqual(typeof(OlePackage), oleObjectAsOlePackage.OleFormat.OlePackage.GetType());
        }

        [Test]
        public void Resize()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertShape(ShapeType.Rectangle, 200, 300);
            // Change shape size and rotation
            shape.Height = 300;
            shape.Width = 500;
            shape.Rotation = 30;

            doc.Save(ArtifactsDir + "Shape.Resize.docx");
        }

        [Test]
        public void LayoutInTableCell()
        {
            //ExStart
            //ExFor:ShapeBase.IsLayoutInCell
            //ExFor:MsWordVersion
            //ExSummary:Shows how to display the shape, inside a table or outside of it.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.StartTable();
            builder.RowFormat.Height = 100;
            builder.RowFormat.HeightRule = HeightRule.Exactly;

            for (int i = 0; i < 31; i++)
            {
                if (i != 0 && i % 7 == 0) builder.EndRow();
                builder.InsertCell();
                builder.Write("Cell contents");
            }

            builder.EndTable();

            NodeCollection runs = doc.GetChildNodes(NodeType.Run, true);
            int num = 1;

            foreach (Run run in runs.OfType<Run>())
            {
                Shape watermark = new Shape(doc, ShapeType.TextPlainText);
                watermark.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
                watermark.RelativeVerticalPosition = RelativeVerticalPosition.Page;
                // False - display the shape outside of table cell, True - display the shape outside of table cell
                watermark.IsLayoutInCell = true; 

                watermark.Width = 30;
                watermark.Height = 30;
                watermark.HorizontalAlignment = HorizontalAlignment.Center;
                watermark.VerticalAlignment = VerticalAlignment.Center;

                watermark.Rotation = -40;
                watermark.Fill.Color = Color.Gainsboro;
                watermark.StrokeColor = Color.Gainsboro;

                watermark.TextPath.Text = string.Format("{0}", num);
                watermark.TextPath.FontFamily = "Arial";

                watermark.Name = $"Watermark_{num++}";
                // Property will take effect only if the WrapType property is set to something other than WrapType.Inline
                watermark.WrapType = WrapType.None; 
                watermark.BehindText = true;

                builder.MoveTo(run);
                builder.InsertNode(watermark);
            }

            // Behavior of Microsoft Word on working with shapes in table cells is changed in the last versions
            // Adding the following line is needed to make the shape displayed in center of a page
            doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2010);

            doc.Save(ArtifactsDir + "Shape.LayoutInTableCell.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.LayoutInTableCell.docx");
            List<Shape> shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().ToList();

            Assert.AreEqual(31, shapes.Count);

            foreach (Shape shape in shapes)
                TestUtil.VerifyShape(ShapeType.TextPlainText, $"Watermark_{shapes.IndexOf(shape) + 1}", 30.0d, 30.0d, 0.0d, 0.0d, shape);
        }

        [Test]
        public void ShapeInsertion()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertShape(ShapeType, RelativeHorizontalPosition, double, RelativeVerticalPosition, double, double, double, WrapType)
            //ExFor:DocumentBuilder.InsertShape(ShapeType, double, double)
            //ExSummary:Shows how to insert DML shapes into the document using a document builder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            // There are two ways of shape insertion
            // These methods allow inserting DML shape into the document model
            // Document must be saved in the format, which supports DML shapes, otherwise, such nodes will be converted
            // to VML shape, while document saving

            // 1. Free-floating shape insertion
            Shape freeFloatingShape = builder.InsertShape(ShapeType.TopCornersRounded, RelativeHorizontalPosition.Page, 100, RelativeVerticalPosition.Page, 100, 50, 50, WrapType.None);
            freeFloatingShape.Rotation = 30.0;
            // 2. Inline shape insertion
            Shape inlineShape = builder.InsertShape(ShapeType.DiagonalCornersRounded, 50, 50);
            inlineShape.Rotation = 30.0;

            // If you need to create "NonPrimitive" shapes, like SingleCornerSnipped, TopCornersSnipped, DiagonalCornersSnipped,
            // TopCornersOneRoundedOneSnipped, SingleCornerRounded, TopCornersRounded, DiagonalCornersRounded
            // please save the document with "Strict" or "Transitional" compliance which allows saving shape as DML
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);
            saveOptions.Compliance = OoxmlCompliance.Iso29500_2008_Transitional;
            
            doc.Save(ArtifactsDir + "Shape.ShapeInsertion.docx", saveOptions);
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.ShapeInsertion.docx");
            List<Shape> shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().ToList();

            TestUtil.VerifyShape(ShapeType.TopCornersRounded, "TopCornersRounded 100002", 50.0d, 50.0d, 100.0d, 100.0d, shapes[0]);
            TestUtil.VerifyShape(ShapeType.DiagonalCornersRounded, "DiagonalCornersRounded 100004", 50.0d, 50.0d, 0.0d, 0.0d, shapes[1]);
        }

        //ExStart
        //ExFor:Shape.Accept(DocumentVisitor)
        //ExFor:Shape.Chart
        //ExFor:Shape.Clone(Boolean, INodeCloningListener)
        //ExFor:Shape.ExtrusionEnabled
        //ExFor:Shape.Filled
        //ExFor:Shape.HasChart
        //ExFor:Shape.OleFormat
        //ExFor:Shape.ShadowEnabled
        //ExFor:Shape.StoryType
        //ExFor:Shape.StrokeColor
        //ExFor:Shape.Stroked
        //ExFor:Shape.StrokeWeight
        //ExSummary:Shows how to iterate over all the shapes in a document.
        [Test] //ExSkip
        public void VisitShapes()
        {
            // Open a document that contains shapes
            Document doc = new Document(MyDir + "Revision shape.docx");
            Assert.AreEqual(2, doc.GetChildNodes(NodeType.Shape, true).Count); //ExSKip

            // Create a ShapeVisitor and get the document to accept it
            ShapeVisitor shapeVisitor = new ShapeVisitor();
            doc.Accept(shapeVisitor);

            // Print all the information that the visitor has collected
            Console.WriteLine(shapeVisitor.GetText());
        }

        /// <summary>
        /// DocumentVisitor implementation that collects information about visited shapes into a StringBuilder, to be printed to the console.
        /// </summary>
        private class ShapeVisitor : DocumentVisitor
        {
            public ShapeVisitor()
            {
                mShapesVisited = 0;
                mTextIndentLevel = 0;
                mStringBuilder = new StringBuilder();
            }

            /// <summary>
            /// Appends a line to the StringBuilder with one prepended tab character for each indent level.
            /// </summary>
            private void AppendLine(string text)
            {
                for (int i = 0; i < mTextIndentLevel; i++) mStringBuilder.Append('\t');

                mStringBuilder.AppendLine(text);
            }

            /// <summary>
            /// Return all the text that the StringBuilder has accumulated.
            /// </summary>
            public string GetText()
            {
                return $"Shapes visited: {mShapesVisited}\n{mStringBuilder}";
            }

            /// <summary>
            /// Called when the start of a Shape node is visited.
            /// </summary>
            public override VisitorAction VisitShapeStart(Shape shape)
            {
                AppendLine($"Shape found: {shape.ShapeType}");

                mTextIndentLevel++;

                if (shape.HasChart)
                    AppendLine($"Has chart: {shape.Chart.Title.Text}");

                AppendLine($"Extrusion enabled: {shape.ExtrusionEnabled}");
                AppendLine($"Shadow enabled: {shape.ShadowEnabled}");
                AppendLine($"StoryType: {shape.StoryType}");

                if (shape.Stroked)
                {
                    Assert.AreEqual(shape.Stroke.Color, shape.StrokeColor);
                    AppendLine($"Stroke colors: {shape.Stroke.Color}, {shape.Stroke.Color2}");
                    AppendLine($"Stroke weight: {shape.StrokeWeight}");

                }

                if (shape.Filled)
                    AppendLine($"Filled: {shape.FillColor}");

                if (shape.OleFormat != null)
                    AppendLine($"Ole found of type: {shape.OleFormat.ProgId}");

                if (shape.SignatureLine != null)
                    AppendLine($"Found signature line for: {shape.SignatureLine.Signer}, {shape.SignatureLine.SignerTitle}");

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when the end of a Shape node is visited.
            /// </summary>
            public override VisitorAction VisitShapeEnd(Shape shape)
            {
                mTextIndentLevel--;
                mShapesVisited++;
                AppendLine($"End of {shape.ShapeType}");

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when the start of a GroupShape node is visited.
            /// </summary>
            public override VisitorAction VisitGroupShapeStart(GroupShape groupShape)
            {
                AppendLine($"Shape group found: {groupShape.ShapeType}");
                mTextIndentLevel++;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when the end of a GroupShape node is visited.
            /// </summary>
            public override VisitorAction VisitGroupShapeEnd(GroupShape groupShape)
            {
                mTextIndentLevel--;
                AppendLine($"End of {groupShape.ShapeType}");

                return VisitorAction.Continue;
            }

            private int mShapesVisited;
            private int mTextIndentLevel;
            private readonly StringBuilder mStringBuilder;
        }
        //ExEnd

        [Test]
        public void SignatureLine()
        {
            //ExStart
            //ExFor:Shape.SignatureLine
            //ExFor:ShapeBase.IsSignatureLine
            //ExFor:SignatureLine
            //ExFor:SignatureLine.AllowComments
            //ExFor:SignatureLine.DefaultInstructions
            //ExFor:SignatureLine.Email
            //ExFor:SignatureLine.Instructions
            //ExFor:SignatureLine.IsSigned
            //ExFor:SignatureLine.IsValid
            //ExFor:SignatureLine.ShowDate
            //ExFor:SignatureLine.Signer
            //ExFor:SignatureLine.SignerTitle
            //ExSummary:Shows how to create a line for a signature and insert it into a document.
            // Create a blank document and its DocumentBuilder
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // The SignatureLineOptions will contain all the data that the signature line will display
            SignatureLineOptions options = new SignatureLineOptions
            {
                AllowComments = true,
                DefaultInstructions = true,
                Email = "john.doe@management.com",
                Instructions = "Please sign here",
                ShowDate = true,
                Signer = "John Doe",
                SignerTitle = "Senior Manager"
            };

            // Insert the signature line, applying our SignatureLineOptions
            // We can control where the signature line will appear on the page using a combination of left/top indents and margin-relative positions
            // Since we are placing the signature line at the bottom right of the page, we will need to use negative indents to move it into view 
            Shape shape = builder.InsertSignatureLine(options, RelativeHorizontalPosition.RightMargin, -170.0, RelativeVerticalPosition.BottomMargin, -60.0, WrapType.None);
            Assert.True(shape.IsSignatureLine);

            // The SignatureLine object is a member of the shape that contains it
            SignatureLine signatureLine = shape.SignatureLine;

            Assert.AreEqual("john.doe@management.com", signatureLine.Email);
            Assert.AreEqual("John Doe", signatureLine.Signer);
            Assert.AreEqual("Senior Manager", signatureLine.SignerTitle);
            Assert.AreEqual("Please sign here", signatureLine.Instructions);
            Assert.True(signatureLine.ShowDate);
            Assert.True(signatureLine.AllowComments);
            Assert.True(signatureLine.DefaultInstructions);

            // We will be prompted to sign it when we open the document
            Assert.False(signatureLine.IsSigned);

            // The object may be valid, but the signature itself isn't until it is signed
            Assert.False(signatureLine.IsValid);

            doc.Save(ArtifactsDir + "Shape.SignatureLine.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.SignatureLine.docx");
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyShape(ShapeType.Image, string.Empty, 192.75d, 96.75d, -60.0d, -170.0d, shape);
            Assert.True(shape.IsSignatureLine);

            signatureLine = shape.SignatureLine;

            Assert.AreEqual("john.doe@management.com", signatureLine.Email);
            Assert.AreEqual("John Doe", signatureLine.Signer);
            Assert.AreEqual("Senior Manager", signatureLine.SignerTitle);
            Assert.AreEqual("Please sign here", signatureLine.Instructions);
            Assert.True(signatureLine.ShowDate);
            Assert.True(signatureLine.AllowComments);
            Assert.True(signatureLine.DefaultInstructions);
            Assert.False(signatureLine.IsSigned);
            Assert.False(signatureLine.IsValid);
        }

        [Test]
        public void TextBox()
        {
            //ExStart
            //ExFor:Shape.TextBox
            //ExFor:Shape.LastParagraph
            //ExFor:TextBox
            //ExFor:TextBox.FitShapeToText
            //ExFor:TextBox.InternalMarginBottom
            //ExFor:TextBox.InternalMarginLeft
            //ExFor:TextBox.InternalMarginRight
            //ExFor:TextBox.InternalMarginTop
            //ExFor:TextBox.LayoutFlow
            //ExFor:TextBox.TextBoxWrapMode
            //ExFor:TextBoxWrapMode
            //ExSummary:Shows how to insert text boxes and arrange their text.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a shape that contains a TextBox
            Shape textBoxShape = builder.InsertShape(ShapeType.TextBox, 150, 100);
            TextBox textBox = textBoxShape.TextBox;

            // Move the document builder to inside the TextBox and write text
            builder.MoveTo(textBoxShape.LastParagraph);
            builder.Write("Vertical text");

            // Text is displayed vertically, written top to bottom
            textBox.LayoutFlow = LayoutFlow.TopToBottomIdeographic;

            // Move the builder out of the shape and back into the main document body
            builder.MoveTo(textBoxShape.ParentParagraph);

            // Insert another TextBox
            textBoxShape = builder.InsertShape(ShapeType.TextBox, 150, 100);
            textBox = textBoxShape.TextBox;

            // Apply these values to both these members to get the parent shape to defy the dimensions we set to fit tightly around the TextBox's text
            textBox.FitShapeToText = true;
            textBox.TextBoxWrapMode = TextBoxWrapMode.None;

            builder.MoveTo(textBoxShape.LastParagraph);
            builder.Write("Text fit tightly inside textbox");

            builder.MoveTo(textBoxShape.ParentParagraph);

            textBoxShape = builder.InsertShape(ShapeType.TextBox, 100, 100);
            textBox = textBoxShape.TextBox;

            // Set margins for the textbox
            textBox.InternalMarginTop = 15;
            textBox.InternalMarginBottom = 15;
            textBox.InternalMarginLeft = 15;
            textBox.InternalMarginRight = 15;

            builder.MoveTo(textBoxShape.LastParagraph);
            builder.Write("Text placed according to textbox margins");

            doc.Save(ArtifactsDir + "Shape.TextBox.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.TextBox.docx");
            List<Shape> shapes = doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>().ToList();

            TestUtil.VerifyShape(ShapeType.TextBox, "TextBox 100002", 150.0d, 100.0d, 0.0d, 0.0d, shapes[0]);
            TestUtil.VerifyTextBox(LayoutFlow.TopToBottomIdeographic, false, TextBoxWrapMode.Square, 3.6d, 3.6d, 7.2d, 7.2d, shapes[0].TextBox);
            Assert.AreEqual("Vertical text", shapes[0].GetText().Trim());

            TestUtil.VerifyShape(ShapeType.TextBox, "TextBox 100004", 150.0d, 100.0d, 0.0d, 0.0d, shapes[1]);
            TestUtil.VerifyTextBox(LayoutFlow.Horizontal, true, TextBoxWrapMode.None, 3.6d, 3.6d, 7.2d, 7.2d, shapes[1].TextBox);
            Assert.AreEqual("Text fit tightly inside textbox", shapes[1].GetText().Trim());

            TestUtil.VerifyShape(ShapeType.TextBox, "TextBox 100006", 100.0d, 100.0d, 0.0d, 0.0d, shapes[2]);
            TestUtil.VerifyTextBox(LayoutFlow.Horizontal, false, TextBoxWrapMode.Square, 15.0d, 15.0d, 15.0d, 15.0d, shapes[2].TextBox);
            Assert.AreEqual("Text placed according to textbox margins", shapes[2].GetText().Trim());
        }

        [Test]
        public void TextBoxShapeType()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set compatibility options to correctly using of VerticalAnchor property
            doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2016);

            Shape textBoxShape = builder.InsertShape(ShapeType.TextBox, 100, 100);
            // Not all formats are compatible with this one
            // For most of incompatible formats AW generated a warnings on save, so use doc.WarningCallback to check it
            textBoxShape.TextBox.VerticalAnchor = TextBoxAnchor.Bottom;
            
            builder.MoveTo(textBoxShape.LastParagraph);
            builder.Write("Text placed bottom");

            doc.Save(ArtifactsDir + "Shape.TextBoxShapeType.docx");
        }

        [Test]
        public void CreateLinkBetweenTextBoxes()
        {
            //ExStart
            //ExFor:TextBox.IsValidLinkTarget(TextBox)
            //ExFor:TextBox.Next
            //ExFor:TextBox.Previous
            //ExFor:TextBox.BreakForwardLink
            //ExSummary:Shows how to work with textbox forward link
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a few textboxes for example
            Shape textBoxShape1 = builder.InsertShape(ShapeType.TextBox, 100, 100);
            TextBox textBox1 = textBoxShape1.TextBox;
            builder.Writeln();
            
            Shape textBoxShape2 = builder.InsertShape(ShapeType.TextBox, 100, 100);
            TextBox textBox2 = textBoxShape2.TextBox;
            builder.Writeln();
            
            Shape textBoxShape3 = builder.InsertShape(ShapeType.TextBox, 100, 100);
            TextBox textBox3 = textBoxShape3.TextBox;
            builder.Writeln();

            Shape textBoxShape4 = builder.InsertShape(ShapeType.TextBox, 100, 100);
            TextBox textBox4 = textBoxShape4.TextBox;
            
            // Create link between textboxes if possible
            if (textBox1.IsValidLinkTarget(textBox2))
                textBox1.Next = textBox2;

            if (textBox2.IsValidLinkTarget(textBox3))
                textBox2.Next = textBox3;

            // You can only create a link on an empty textbox
            builder.MoveTo(textBoxShape4.LastParagraph);
            builder.Write("Vertical text");

            // Thus, this textbox is not a valid link target
            Assert.IsFalse(textBox3.IsValidLinkTarget(textBox4));
            
            if (textBox1.Next != null && textBox1.Previous == null)
                Console.WriteLine("This TextBox is the head of the sequence");
 
            if (textBox2.Next != null && textBox2.Previous != null)
                Console.WriteLine("This TextBox is the middle of the sequence");
 
            if (textBox3.Next == null && textBox3.Previous != null)
            {
                Console.WriteLine("This TextBox is the tail of the sequence");
                
                // Break the forward link between textBox2 and textBox3
                textBox3.Previous.BreakForwardLink();
                // Check that link was break successfully
                Assert.IsTrue(textBox2.Next == null);
                Assert.IsTrue(textBox3.Previous == null);
            }

            doc.Save(ArtifactsDir + "Shape.CreateLinkBetweenTextBoxes.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.CreateLinkBetweenTextBoxes.docx");
            List<Shape> shapes = doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>().ToList();

            TestUtil.VerifyShape(ShapeType.TextBox, "TextBox 100002", 100.0d, 100.0d, 0.0d, 0.0d, shapes[0]);
            TestUtil.VerifyTextBox(LayoutFlow.Horizontal, false, TextBoxWrapMode.Square, 3.6d, 3.6d, 7.2d, 7.2d, shapes[0].TextBox);
            Assert.AreEqual(string.Empty, shapes[0].GetText().Trim());

            TestUtil.VerifyShape(ShapeType.TextBox, "TextBox 100004", 100.0d, 100.0d, 0.0d, 0.0d, shapes[1]);
            TestUtil.VerifyTextBox(LayoutFlow.Horizontal, false, TextBoxWrapMode.Square, 3.6d, 3.6d, 7.2d, 7.2d, shapes[1].TextBox);
            Assert.AreEqual(string.Empty, shapes[1].GetText().Trim());

            TestUtil.VerifyShape(ShapeType.Rectangle, "TextBox 100006", 100.0d, 100.0d, 0.0d, 0.0d, shapes[2]);
            TestUtil.VerifyTextBox(LayoutFlow.Horizontal, false, TextBoxWrapMode.Square, 3.6d, 3.6d, 7.2d, 7.2d, shapes[2].TextBox);
            Assert.AreEqual(string.Empty, shapes[2].GetText().Trim());

            TestUtil.VerifyShape(ShapeType.TextBox, "TextBox 100008", 100.0d, 100.0d, 0.0d, 0.0d, shapes[3]);
            TestUtil.VerifyTextBox(LayoutFlow.Horizontal, false, TextBoxWrapMode.Square, 3.6d, 3.6d, 7.2d, 7.2d, shapes[3].TextBox);
            Assert.AreEqual("Vertical text", shapes[3].GetText().Trim());
        }

        [Test]
        public void GetTextBoxAndChangeTextAnchor()
        {
            //ExStart
            //ExFor:TextBoxAnchor
            //ExFor:TextBox.VerticalAnchor
            //ExSummary:Shows how to change text position inside textbox shape.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 200);
            textBox.TextBox.VerticalAnchor = TextBoxAnchor.Bottom;
            
            builder.MoveTo(textBox.FirstParagraph);
            builder.Write("Textbox contents");

            doc.Save(ArtifactsDir + "Shape.GetTextBoxAndChangeAnchor.docx");
            //ExEnd
            
            doc = new Document(ArtifactsDir + "Shape.GetTextBoxAndChangeAnchor.docx");
            textBox = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyShape(ShapeType.TextBox, "TextBox 100002", 200.0d, 200.0d, 0.0d, 0.0d, textBox);
            TestUtil.VerifyTextBox(LayoutFlow.Horizontal, false, TextBoxWrapMode.Square, 3.6d, 3.6d, 7.2d, 7.2d, textBox.TextBox);
            Assert.AreEqual("Textbox contents", textBox.GetText().Trim());
        }

        //ExStart
        //ExFor:Shape.TextPath
        //ExFor:ShapeBase.IsWordArt
        //ExFor:TextPath
        //ExFor:TextPath.Bold
        //ExFor:TextPath.FitPath
        //ExFor:TextPath.FitShape
        //ExFor:TextPath.FontFamily
        //ExFor:TextPath.Italic
        //ExFor:TextPath.Kerning
        //ExFor:TextPath.On
        //ExFor:TextPath.ReverseRows
        //ExFor:TextPath.RotateLetters
        //ExFor:TextPath.SameLetterHeights
        //ExFor:TextPath.Shadow
        //ExFor:TextPath.SmallCaps
        //ExFor:TextPath.Spacing
        //ExFor:TextPath.StrikeThrough
        //ExFor:TextPath.Text
        //ExFor:TextPath.TextPathAlignment
        //ExFor:TextPath.Trim
        //ExFor:TextPath.Underline
        //ExFor:TextPath.XScale
        //ExFor:TextPathAlignment
        //ExSummary:Shows how to work with WordArt.
        [Test] //ExSkip
        public void InsertTextPaths()
        {
            Document doc = new Document();

            // Insert a WordArt object and capture the shape that contains it in a variable
            Shape shape = AppendWordArt(doc, "Bold & Italic", "Arial", 240, 24, Color.White, Color.Black, ShapeType.TextPlainText);

            // View and verify various text formatting settings
            shape.TextPath.Bold = true;
            shape.TextPath.Italic = true;

            Assert.False(shape.TextPath.Underline);
            Assert.False(shape.TextPath.Shadow);
            Assert.False(shape.TextPath.StrikeThrough);
            Assert.False(shape.TextPath.ReverseRows);
            Assert.False(shape.TextPath.XScale);
            Assert.False(shape.TextPath.Trim);
            Assert.False(shape.TextPath.SmallCaps);

            Assert.AreEqual(36.0, shape.TextPath.Size);
            Assert.AreEqual("Bold & Italic", shape.TextPath.Text);
            Assert.AreEqual(ShapeType.TextPlainText, shape.ShapeType);

            // Toggle whether to display text
            shape = AppendWordArt(doc, "On set to true", "Calibri", 150, 24, Color.Yellow, Color.Red, ShapeType.TextPlainText);
            shape.TextPath.On = true;

            shape = AppendWordArt(doc, "On set to false", "Calibri", 150, 24, Color.Yellow, Color.Purple, ShapeType.TextPlainText);
            shape.TextPath.On = false;

            // Apply kerning
            shape = AppendWordArt(doc, "Kerning: VAV", "Times New Roman", 90, 24, Color.Orange, Color.Red, ShapeType.TextPlainText);
            shape.TextPath.Kerning = true;

            shape = AppendWordArt(doc, "No kerning: VAV", "Times New Roman", 100, 24, Color.Orange, Color.Red, ShapeType.TextPlainText);
            shape.TextPath.Kerning = false;

            // Apply custom spacing, on a scale from 0.0 (none) to 1.0 (default)
            shape = AppendWordArt(doc, "Spacing set to 0.1", "Calibri", 120, 24, Color.BlueViolet, Color.Blue, ShapeType.TextCascadeDown);
            shape.TextPath.Spacing = 0.1;

            // Rotate letters 90 degrees to the left, text is still laid out horizontally
            shape = AppendWordArt(doc, "RotateLetters", "Calibri", 200, 36, Color.GreenYellow, Color.Green, ShapeType.TextWave);
            shape.TextPath.RotateLetters = true;

            // Set the x-height to equal the cap height
            shape = AppendWordArt(doc, "Same character height for lower and UPPER case", "Calibri", 300, 24, Color.DeepSkyBlue, Color.DodgerBlue, ShapeType.TextSlantUp);
            shape.TextPath.SameLetterHeights = true;

            // By default, the size of the text will scale to always fit the size of the containing shape, overriding the text size setting
            shape = AppendWordArt(doc, "FitShape on", "Calibri", 160, 24, Color.LightBlue, Color.Blue, ShapeType.TextPlainText);
            Assert.True(shape.TextPath.FitShape);
            shape.TextPath.Size = 24.0;

            // If we set FitShape to false, the size of the text will defy the shape bounds and always keep the size value we set below
            // We can also set TextPathAlignment to align the text
            shape = AppendWordArt(doc, "FitShape off", "Calibri", 160, 24, Color.LightBlue, Color.Blue, ShapeType.TextPlainText);
            shape.TextPath.FitShape = false;
            shape.TextPath.Size = 24.0;
            shape.TextPath.TextPathAlignment = TextPathAlignment.Right;

            doc.Save(ArtifactsDir + "Shape.InsertTextPaths.docx");
            TestInsertTextPaths(ArtifactsDir + "Shape.InsertTextPaths.docx"); //ExSkip
        }

        /// <summary>
        /// Insert a new paragraph with a WordArt shape inside it.
        /// </summary>
        private static Shape AppendWordArt(Document doc, string text, string textFontFamily, double shapeWidth, double shapeHeight, Color wordArtFill, Color line, ShapeType wordArtShapeType)
        {
            // Insert a new paragraph
            Paragraph para = (Paragraph)doc.FirstSection.Body.AppendChild(new Paragraph(doc));

            // Create an inline Shape, which will serve as a container for our WordArt, and append it to the paragraph
            // The shape can only be a valid WordArt shape if the ShapeType assigned here is a WordArt-designated ShapeType
            // These types will have "WordArt object" in the description and their enumerator names will start with "Text..."
            Shape shape = new Shape(doc, wordArtShapeType);
            shape.WrapType = WrapType.Inline;
            para.AppendChild(shape);

            // Set the shape's width and height
            shape.Width = shapeWidth;
            shape.Height = shapeHeight;

            // These color settings will apply to the letters of the displayed WordArt text
            shape.FillColor = wordArtFill;
            shape.StrokeColor = line;

            // The WordArt object is accessed here, and we will set the text and font like this
            shape.TextPath.Text = text;
            shape.TextPath.FontFamily = textFontFamily;
            
            return shape;
        }
        //ExEnd

        private void TestInsertTextPaths(string filename)
        {
            Document doc = new Document(filename);
            List<Shape> shapes = doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>().ToList();

            TestUtil.VerifyShape(ShapeType.TextPlainText, string.Empty, 240, 24, 0.0d, 0.0d, shapes[0]);
            Assert.True(shapes[0].TextPath.Bold);
            Assert.True(shapes[0].TextPath.Italic);

            TestUtil.VerifyShape(ShapeType.TextPlainText, string.Empty, 150, 24, 0.0d, 0.0d, shapes[1]);
            Assert.True(shapes[1].TextPath.On);

            TestUtil.VerifyShape(ShapeType.TextPlainText, string.Empty, 150, 24, 0.0d, 0.0d, shapes[2]);
            Assert.False(shapes[2].TextPath.On);

            TestUtil.VerifyShape(ShapeType.TextPlainText, string.Empty, 90, 24, 0.0d, 0.0d, shapes[3]);
            Assert.True(shapes[3].TextPath.Kerning);

            TestUtil.VerifyShape(ShapeType.TextPlainText, string.Empty, 100, 24, 0.0d, 0.0d, shapes[4]);
            Assert.False(shapes[4].TextPath.Kerning);

            TestUtil.VerifyShape(ShapeType.TextCascadeDown, string.Empty, 120, 24, 0.0d, 0.0d, shapes[5]);
            Assert.AreEqual(0.1d, shapes[5].TextPath.Spacing, 0.01d);

            TestUtil.VerifyShape(ShapeType.TextWave, string.Empty, 200, 36, 0.0d, 0.0d, shapes[6]);
            Assert.True(shapes[6].TextPath.RotateLetters);

            TestUtil.VerifyShape(ShapeType.TextSlantUp, string.Empty, 300, 24, 0.0d, 0.0d, shapes[7]);
            Assert.True(shapes[7].TextPath.SameLetterHeights);

            TestUtil.VerifyShape(ShapeType.TextPlainText, string.Empty, 160, 24, 0.0d, 0.0d, shapes[8]);
            Assert.True(shapes[8].TextPath.FitShape);
            Assert.AreEqual(24.0d, shapes[8].TextPath.Size);

            TestUtil.VerifyShape(ShapeType.TextPlainText, string.Empty, 160, 24, 0.0d, 0.0d, shapes[9]);
            Assert.False(shapes[9].TextPath.FitShape);
            Assert.AreEqual(24.0d, shapes[9].TextPath.Size);
            Assert.AreEqual(TextPathAlignment.Right, shapes[9].TextPath.TextPathAlignment);
        }

        [Test]
        public void ShapeRevision()
        {
            //ExStart
            //ExFor:ShapeBase.IsDeleteRevision
            //ExFor:ShapeBase.IsInsertRevision
            //ExSummary:Shows how to work with revision shapes.
            // Open a blank document
            Document doc = new Document();

            // Insert an inline shape without tracking revisions
            Assert.False(doc.TrackRevisions);
            Shape shape = new Shape(doc, ShapeType.Cube);
            shape.WrapType = WrapType.Inline;
            shape.Width = 100.0;
            shape.Height = 100.0;
            doc.FirstSection.Body.FirstParagraph.AppendChild(shape);

            // Start tracking revisions and then insert another shape
            doc.StartTrackRevisions("John Doe");

            shape = new Shape(doc, ShapeType.Sun);
            shape.WrapType = WrapType.Inline;
            shape.Width = 100.0;
            shape.Height = 100.0;
            doc.FirstSection.Body.FirstParagraph.AppendChild(shape);

            // Get the document's shape collection which includes just the two shapes we added
            List<Shape> shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().ToList();
            Assert.AreEqual(2, shapes.Count);

            // Remove the first shape
            shapes[0].Remove();

            // Because we removed that shape while changes were being tracked, the shape counts as a delete revision
            Assert.AreEqual(ShapeType.Cube, shapes[0].ShapeType);
            Assert.True(shapes[0].IsDeleteRevision);

            // And we inserted another shape while tracking changes, so that shape will count as an insert revision
            Assert.AreEqual(ShapeType.Sun, shapes[1].ShapeType);
            Assert.True(shapes[1].IsInsertRevision);
            //ExEnd
        }

        [Test]
        public void MoveRevisions()
        {
            //ExStart
            //ExFor:ShapeBase.IsMoveFromRevision
            //ExFor:ShapeBase.IsMoveToRevision
            //ExSummary:Shows how to identify move revision shapes.
            // Open a document that contains a move revision
            // A move revision is when we, while changes are tracked, cut(not copy)-and-paste or highlight and drag text from one place to another
            // If inline shapes are caught up in the text movement, they will count as move revisions as well
            // Moving a floating shape will not count as a move revision
            Document doc = new Document(MyDir + "Revision shape.docx");

            // The document has one shape that was moved, but shape move revisions will have two instances of that shape
            // One will be the shape at its arrival destination and the other will be the shape at its original location
            List<Shape> nc = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().ToList();
            Assert.AreEqual(2, nc.Count);

            // This is the move to revision, also the shape at its arrival destination
            Assert.False(nc[0].IsMoveFromRevision);
            Assert.True(nc[0].IsMoveToRevision);

            // This is the move from revision, which is the shape at its original location
            Assert.True(nc[1].IsMoveFromRevision);
            Assert.False(nc[1].IsMoveToRevision);
            //ExEnd
        }

        [Test]
        public void AdjustWithEffects()
        {
            //ExStart
            //ExFor:ShapeBase.AdjustWithEffects(RectangleF)
            //ExFor:ShapeBase.BoundsWithEffects
            //ExSummary:Shows how to check how a shape's bounds are affected by shape effects.
            // Open a document that contains two shapes and get its shape collection
            Document doc = new Document(MyDir + "Shape shadow effect.docx");
            List<Shape> shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().ToList();
            Assert.AreEqual(2, shapes.Count);

            // The two shapes are identical in terms of dimensions and shape type
            Assert.AreEqual(shapes[0].Width, shapes[1].Width);
            Assert.AreEqual(shapes[0].Height, shapes[1].Height);
            Assert.AreEqual(shapes[0].ShapeType, shapes[1].ShapeType);

            // However, the first shape has no effects, while the second one has a shadow and thick outline
            Assert.AreEqual(0.0, shapes[0].StrokeWeight);
            Assert.AreEqual(20.0, shapes[1].StrokeWeight);
            Assert.False(shapes[0].ShadowEnabled);
            Assert.True(shapes[1].ShadowEnabled);

            // These effects make the size of the second shape's silhouette bigger than that of the first
            // Even though the size of the rectangle that shows up when we click on these shapes in Microsoft Word is the same,
            // the practical outer bounds of the second shape are affected by the shadow and outline and are bigger
            // We can use the AdjustWithEffects method to see exactly how much bigger they are

            // The first shape has no outline or effects
            Shape shape = shapes[0];

            // Create a RectangleF object, which represents a rectangle, which we could potentially use as the coordinates and bounds for a shape
            RectangleF rectangleF = new RectangleF(200, 200, 1000, 1000);

            // Run this method to get the size of the rectangle adjusted for all our shape's effects
            RectangleF rectangleFOut = shape.AdjustWithEffects(rectangleF);

            // Since the shape has no border-changing effects, its boundary dimensions are unaffected
            Assert.AreEqual(200, rectangleFOut.X);
            Assert.AreEqual(200, rectangleFOut.Y);
            Assert.AreEqual(1000, rectangleFOut.Width);
            Assert.AreEqual(1000, rectangleFOut.Height);

            // The final extent of the first shape, in points
            Assert.AreEqual(0, shape.BoundsWithEffects.X);
            Assert.AreEqual(0, shape.BoundsWithEffects.Y);
            Assert.AreEqual(147, shape.BoundsWithEffects.Width);
            Assert.AreEqual(147, shape.BoundsWithEffects.Height);

            // Do the same with the second shape
            shape = shapes[1];
            rectangleF = new RectangleF(200, 200, 1000, 1000);
            rectangleFOut = shape.AdjustWithEffects(rectangleF);
            
            // The shape's x/y coordinates (top left corner location) have been pushed back by the thick outline
            Assert.AreEqual(171.5, rectangleFOut.X);
            Assert.AreEqual(167, rectangleFOut.Y);

            // The width and height were also affected by the outline and shadow
            Assert.AreEqual(1045, rectangleFOut.Width);
            Assert.AreEqual(1132, rectangleFOut.Height);

            // These values are also affected by effects
            Assert.AreEqual(-28.5, shape.BoundsWithEffects.X);
            Assert.AreEqual(-33, shape.BoundsWithEffects.Y);
            Assert.AreEqual(192, shape.BoundsWithEffects.Width);
            Assert.AreEqual(279, shape.BoundsWithEffects.Height);
            //ExEnd
        }

        [Test]
        public void RenderAllShapes()
        {
            //ExStart
            //ExFor:ShapeBase.GetShapeRenderer
            //ExFor:NodeRendererBase.Save(Stream, ImageSaveOptions)
            //ExSummary:Shows how to export shapes to files in the local file system using a shape renderer.
            // Open a document that contains shapes and get its shape collection
            Document doc = new Document(MyDir + "Various shapes.docx");
            List<Shape> shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().ToList();
            Assert.AreEqual(7, shapes.Count);

            // There are 7 shapes in the document, with one group shape with 2 child shapes
            // The child shapes will be rendered but their parent group shape will be skipped, so we will see 6 output files
            foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>())
            {
                ShapeRenderer renderer = shape.GetShapeRenderer();
                ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);
                renderer.Save(ArtifactsDir + $"Shape.RenderAllShapes.{shape.Name}.png", options);
            }
            //ExEnd
        }

        [Test]
        public void DocumentHasSmartArtObject()
        {
            //ExStart
            //ExFor:Shape.HasSmartArt
            //ExSummary:Shows how to detect that Shape has a SmartArt object.
            Document doc = new Document(MyDir + "SmartArt.docx");
 
            int count = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().Count(shape => shape.HasSmartArt);

            Console.WriteLine("The document has {0} shapes with SmartArt.", count);
            //ExEnd

            Assert.AreEqual(2, count);
        }

        [Test, Category("SkipMono")]
        public void OfficeMathRenderer()
        {
            //ExStart
            //ExFor:NodeRendererBase
            //ExFor:NodeRendererBase.BoundsInPoints
            //ExFor:NodeRendererBase.GetBoundsInPixels(Single, Single)
            //ExFor:NodeRendererBase.GetBoundsInPixels(Single, Single, Single)
            //ExFor:NodeRendererBase.GetOpaqueBoundsInPixels(Single, Single)
            //ExFor:NodeRendererBase.GetOpaqueBoundsInPixels(Single, Single, Single)
            //ExFor:NodeRendererBase.GetSizeInPixels(Single, Single)
            //ExFor:NodeRendererBase.GetSizeInPixels(Single, Single, Single)
            //ExFor:NodeRendererBase.OpaqueBoundsInPoints
            //ExFor:NodeRendererBase.SizeInPoints
            //ExFor:OfficeMathRenderer
            //ExFor:OfficeMathRenderer.#ctor(Math.OfficeMath)
            //ExSummary:Shows how to measure and scale shapes.
            // Open a document that contains an OfficeMath object
            Document doc = new Document(MyDir + "Office math.docx");

            // Create a renderer for the OfficeMath object 
            OfficeMath officeMath = (OfficeMath)doc.GetChild(NodeType.OfficeMath, 0, true);
            OfficeMathRenderer renderer = new OfficeMathRenderer(officeMath);

            // We can measure the size of the image that the OfficeMath object will create when we render it
            Assert.AreEqual(119.0f, renderer.SizeInPoints.Width, 0.2f);
            Assert.AreEqual(13.0f, renderer.SizeInPoints.Height, 0.1f);

            Assert.AreEqual(119.0f, renderer.BoundsInPoints.Width, 0.2f);
            Assert.AreEqual(13.0f, renderer.BoundsInPoints.Height, 0.1f);

            // Shapes with transparent parts may return different values here
            Assert.AreEqual(119.0f, renderer.OpaqueBoundsInPoints.Width, 0.2f);
            Assert.AreEqual(14.2f, renderer.OpaqueBoundsInPoints.Height, 0.1f);

            // Get the shape size in pixels, with linear scaling to a specific DPI
            Rectangle bounds = renderer.GetBoundsInPixels(1.0f, 96.0f);
            Assert.AreEqual(159, bounds.Width);
            Assert.AreEqual(18, bounds.Height);

            // Get the shape size in pixels, but with a different DPI for the horizontal and vertical dimensions
            bounds = renderer.GetBoundsInPixels(1.0f, 96.0f, 150.0f);
            Assert.AreEqual(159, bounds.Width);
            Assert.AreEqual(28, bounds.Height);

            // The opaque bounds may vary here also
            bounds = renderer.GetOpaqueBoundsInPixels(1.0f, 96.0f);
            Assert.AreEqual(159, bounds.Width);
            Assert.AreEqual(18, bounds.Height);

            bounds = renderer.GetOpaqueBoundsInPixels(1.0f, 96.0f, 150.0f);
            Assert.AreEqual(159, bounds.Width);
            Assert.AreEqual(30, bounds.Height);
            //ExEnd
        }
    }
}