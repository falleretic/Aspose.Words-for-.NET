﻿using Aspose.Words.WebExtensions;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Web_Extensions
{
    class WorkingWithWebExtension : TestDataHelper
    {
        public static void Run()
        {
            UsingWebExtensionTaskPanes();
        }

        public static void UsingWebExtensionTaskPanes()
        {
            //ExStart:UsingWebExtensionTaskPanes
            Document doc = new Document();

            TaskPane taskPane = new TaskPane();
            doc.WebExtensionTaskPanes.Add(taskPane);

            taskPane.DockState = TaskPaneDockState.Right;
            taskPane.IsVisible = true;
            taskPane.Width = 300;

            taskPane.WebExtension.Reference.Id = "wa102923726";
            taskPane.WebExtension.Reference.Version = "1.0.0.0";
            taskPane.WebExtension.Reference.StoreType = WebExtensionStoreType.OMEX;
            taskPane.WebExtension.Reference.Store = "th-TH";
            taskPane.WebExtension.Properties.Add(new WebExtensionProperty("mailchimpCampaign", "mailchimpCampaign"));
            taskPane.WebExtension.Bindings.Add(new WebExtensionBinding("UnnamedBinding_0_1506535429545",
                WebExtensionBindingType.Text, "194740422"));

            doc.Save(ArtifactsDir + "output.docx");
            //ExEnd:UsingWebExtensionTaskPanes
        }
    }
}