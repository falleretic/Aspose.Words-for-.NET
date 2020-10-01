using System;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_with_Documents.Document_Content
{
    class VbaMacros : TestDataHelper
    {
        [Test]
        public static void CreateVbaProject()
        {
            //ExStart:CreateVbaProject
            Document doc = new Document();

            // Create a new VBA project
            VbaProject project = new VbaProject();
            project.Name = "AsposeProject";
            doc.VbaProject = project;

            // Create a new module and specify a macro source code
            VbaModule module = new VbaModule();
            module.Name = "AsposeModule";
            module.Type = VbaModuleType.ProceduralModule;
            module.SourceCode = "New source code";

            // Add module to the VBA project
            doc.VbaProject.Modules.Add(module);

            doc.Save(ArtifactsDir + "VbaMacros.CreateVbaProject.docm");
            //ExEnd:CreateVbaProject
        }

        [Test]
        public static void ReadVbaMacros()
        {
            //ExStart:ReadVbaMacros
            Document doc = new Document(MyDir + "VbaMacros.CreateVbaProject.docm");

            if (doc.VbaProject != null)
            {
                foreach (VbaModule module in doc.VbaProject.Modules)
                {
                    Console.WriteLine(module.SourceCode);
                }
            }
            //ExEnd:ReadVbaMacros
        }

        [Test]
        public static void ModifyVbaMacros()
        {
            //ExStart:ModifyVbaMacros
            Document doc = new Document(MyDir + "VbaMacros.CreateVbaProject.docm");
            VbaProject project = doc.VbaProject;

            const string newSourceCode = "Test change source code";

            // Choose a module, and set a new source code
            project.Modules[0].SourceCode = newSourceCode;
            //ExEnd:ModifyVbaMacros


            doc.Save(ArtifactsDir + "VbaProject_out.docm");
            //ExEnd:ModifyVbaMacros
        }

        [Test]
        public static void CloneVbaProject()
        {
            //ExStart:CloneVbaProject
            Document doc = new Document(MyDir + "VbaMacros.CreateVbaProject.docm");
            Document destDoc = new Document();

            // Clone the whole project
            destDoc.VbaProject = doc.VbaProject.Clone();

            destDoc.Save(ArtifactsDir + "output.docm");
            //ExEnd:CloneVbaProject
        }

        [Test]
        public static void CloneVbaModule()
        {
            //ExStart:CloneVbaModule
            Document doc = new Document(MyDir + "VbaMacros.CreateVbaProject.docm");
            Document destDoc = new Document();

            destDoc.VbaProject = new VbaProject();
            // Clone a single module
            VbaModule copyModule = doc.VbaProject.Modules["AsposeModule"].Clone();
            destDoc.VbaProject.Modules.Add(copyModule);

            destDoc.Save(ArtifactsDir + "output.docm");
            //ExEnd:CloneVbaModule
        }
    }
}