using System;
using System.Linq;
using Aspose.Words;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents
{
    internal class WorkingWithVba : DocsExamplesBase
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

            doc.Save(ArtifactsDir + "WorkingWithVba.CreateVbaProject.docm");
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
            
            doc.Save(ArtifactsDir + "WorkingWithVba.ModifyVbaMacros.docm");
            //ExEnd:ModifyVbaMacros
        }

        [Test]
        public static void CloneVbaProject()
        {
            //ExStart:CloneVbaProject
            Document doc = new Document(MyDir + "VBA project.docm");
            Document destDoc = new Document { VbaProject = doc.VbaProject.Clone() };

            destDoc.Save(ArtifactsDir + "WorkingWithVba.CloneVbaProject.docm");
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

            destDoc.Save(ArtifactsDir + "WorkingWithVba.CloneVbaModule.docm");
            //ExEnd:CloneVbaModule
        }

        [Test]
        public static void RemoveBrokenRef()
        {
            //ExStart:RemoveReferenceFromCollectionOfReferences
            Document doc = new Document(MyDir + "VbaProject.docm");

            // Find and remove the reference with some LibId path.
            const string brokenPath = "brokenPath.dll";
            VbaReferenceCollection references = doc.VbaProject.References;
            for (int i = references.Count - 1; i >= 0; i--)
            {
                VbaReference reference = doc.VbaProject.References.ElementAt(i);

                string path = GetLibIdPath(reference);
                if (path == brokenPath)
                    references.RemoveAt(i);
            }

            doc.Save(ArtifactsDir + "WorkingWithVba.RemoveBrokenRef.docm");
            //ExEnd:RemoveReferenceFromCollectionOfReferences
        }
        //ExStart:GetLibIdAndReferencePath
        /// <summary>
        /// Returns string representing LibId path of a specified reference. 
        /// </summary>
        private static string GetLibIdPath(VbaReference reference)
        {
            switch (reference.Type)
            {
                case VbaReferenceType.Registered:
                case VbaReferenceType.Original:
                case VbaReferenceType.Control:
                    return GetLibIdReferencePath(reference.LibId);
                case VbaReferenceType.Project:
                    return GetLibIdProjectPath(reference.LibId);
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        /// <summary>
        /// Returns path from a specified identifier of an Automation type library.
        /// </summary>
        /// <remarks>
        /// Please see details for the syntax at [MS-OVBA], 2.1.1.8 LibidReference. 
        /// </remarks>
        private static string GetLibIdReferencePath(string libIdReference)
        {
            if (libIdReference != null)
            {
                string[] refParts = libIdReference.Split('#');
                if (refParts.Length > 3)
                    return refParts[3];
            }

            return "";
        }

        /// <summary>
        /// Returns path from a specified identifier of an Automation type library.
        /// </summary>
        /// <remarks>
        /// Please see details for the syntax at [MS-OVBA], 2.1.1.12 ProjectReference. 
        /// </remarks>
        private static string GetLibIdProjectPath(string libIdProject)
        {
            return (libIdProject != null) ? libIdProject.Substring(3) : "";
        }
        //ExEnd:GetLibIdAndReferencePath
    }
}