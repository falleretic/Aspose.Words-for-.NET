using System;
using System.Net;
using Aspose.Words.Loading;

namespace Aspose.Words.Examples.CSharp
{
    //ExStart:ImageLoadingWithCredentialsHandler
    public class ImageLoadingWithCredentialsHandler : IResourceLoadingCallback
    {
        public ImageLoadingWithCredentialsHandler()
        {
            mWebClient = new WebClient();
        }

        public ResourceLoadingAction ResourceLoading(ResourceLoadingArgs args)
        {
            if (args.ResourceType == ResourceType.Image)
            {
                Uri uri = new Uri(args.Uri);

                mWebClient.Credentials = uri.Host == "www.aspose.com"
                    ? new NetworkCredential("User1", "akjdlsfkjs")
                    : new NetworkCredential("SomeOtherUserID", "wiurlnlvs");

                // Download the bytes from the location referenced by the URI
                byte[] imageBytes = mWebClient.DownloadData(args.Uri);

                args.SetData(imageBytes);

                return ResourceLoadingAction.UserProvided;
            }

            return ResourceLoadingAction.Default;
        }

        private readonly WebClient mWebClient;
    }
    //ExEnd:ImageLoadingWithCredentialsHandler
}