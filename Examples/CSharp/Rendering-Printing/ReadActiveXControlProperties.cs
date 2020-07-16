using System;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Ole;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class ReadActiveXControlProperties : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            Document doc = new Document(RenderingPrintingDir + "ActiveXControl.docx");

            string properties = "";
            // Retrieve shapes from the document
            foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
            {
                if (shape.OleFormat is null) break;

                OleControl oleControl = shape.OleFormat.OleControl;
                if (oleControl.IsForms2OleControl)
                {
                    Forms2OleControl checkBox = (Forms2OleControl) oleControl;
                    properties = properties + "\nCaption: " + checkBox.Caption;
                    properties = properties + "\nValue: " + checkBox.Value;
                    properties = properties + "\nEnabled: " + checkBox.Enabled;
                    properties = properties + "\nType: " + checkBox.Type;
                    if (checkBox.ChildNodes != null)
                    {
                        properties = properties + "\nChildNodes: " + checkBox.ChildNodes;
                    }

                    properties += "\n";
                }
            }

            properties = properties + "\nTotal ActiveX Controls found: " + doc.GetChildNodes(NodeType.Shape, true).Count;
            Console.WriteLine("\n" + properties);
        }
    }
}