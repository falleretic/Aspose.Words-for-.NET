using System;
using Aspose.Words;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents.Document_Content
{
    internal class WorkingWithVbaMacros : DocsExamplesBase
    {
        [Test]
        public static void CreateVbaProject()
        {
            //ExStart:CreateVbaProject
            Document doc = new Document();

            VbaProject project = new VbaProject();
            project.Name = "AsposeProject";
            doc.VbaProject = project;

            // Create a new module and specify a macro source code.
            VbaModule module = new VbaModule();
            module.Name = "AsposeModule";
            module.Type = VbaModuleType.ProceduralModule;
            module.SourceCode = "New source code";

            // Add module to the VBA project.
            doc.VbaProject.Modules.Add(module);

            doc.Save(ArtifactsDir + "WorkingWithVbaMacros.CreateVbaProject.docm");
            //ExEnd:CreateVbaProject
        }

        [Test]
        public static void ReadVbaMacros()
        {
            //ExStart:ReadVbaMacros
            Document doc = new Document(MyDir + "VBA project.docm");

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
            Document doc = new Document(MyDir + "VBA project.docm");

            VbaProject project = doc.VbaProject;

            const string newSourceCode = "Test change source code";
            project.Modules[0].SourceCode = newSourceCode;
            //ExEnd:ModifyVbaMacros
            
            doc.Save(ArtifactsDir + "WorkingWithVbaMacros.ModifyVbaMacros.docm");
            //ExEnd:ModifyVbaMacros
        }

        [Test]
        public static void CloneVbaProject()
        {
            //ExStart:CloneVbaProject
            Document doc = new Document(MyDir + "VBA project.docm");
            Document destDoc = new Document { VbaProject = doc.VbaProject.Clone() };

            destDoc.Save(ArtifactsDir + "WorkingWithVbaMacros.CloneVbaProject.docm");
            //ExEnd:CloneVbaProject
        }

        [Test]
        public static void CloneVbaModule()
        {
            //ExStart:CloneVbaModule
            Document doc = new Document(MyDir + "VBA project.docm");
            Document destDoc = new Document { VbaProject = new VbaProject() };
            
            VbaModule copyModule = doc.VbaProject.Modules["Module1"].Clone();
            destDoc.VbaProject.Modules.Add(copyModule);

            destDoc.Save(ArtifactsDir + "WorkingWithVbaMacros.CloneVbaModule.docm");
            //ExEnd:CloneVbaModule
        }
    }
}