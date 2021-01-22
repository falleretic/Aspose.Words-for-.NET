﻿// Copyright (c) Aspose 2002-2021. All Rights Reserved.

/*
    This project uses NuGet's Automatic Package Restore feature to 
    resolve the Aspose.Words for .NET API reference when the project is built. 
    Please visit https://docs.nuget.org/consume/nuget-faq for more information. 

    If you do not wish to use NuGet, you can manually download Aspose.Words for .NET API 
    from http://www.aspose.com/downloads, install it, and then add a reference to it to this project. 

    For any issues, questions or suggestions, please visit the Aspose Forums: https://forum.aspose.com/
*/

using Aspose.Words;

namespace Aspose.Plugins.AsposeVSOpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            string FilePath = @"..\..\..\..\Sample Files\";
            string fileName = FilePath + "Remove Hidden Text.docx";

            Document doc = new Document(fileName);
            foreach (Paragraph par in doc.GetChildNodes(NodeType.Paragraph, true))
            {
                par.ParagraphBreakFont.Hidden = false;
                foreach (Run run in par.GetChildNodes(NodeType.Run, true))
                {
                    if (run.Font.Hidden)
                        run.Font.Hidden = false;
                }
            }
            doc.Save(fileName);
        }
    }
}
