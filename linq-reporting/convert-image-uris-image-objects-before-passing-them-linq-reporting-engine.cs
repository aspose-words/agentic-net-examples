using System;
using System.IO;
using System.Net.Http;
using Aspose.Words;
using Aspose.Words.Loading;

public class ImageUriHandler : IResourceLoadingCallback
{
    // This callback intercepts every external resource load request.
    // For image resources we download the image data and supply it to Aspose.Words.
    public ResourceLoadingAction ResourceLoading(ResourceLoadingArgs args)
    {
        if (args.ResourceType == ResourceType.Image)
        {
            // The OriginalUri property contains the URI that was specified in the document.
            string uri = args.OriginalUri;

            // Simple heuristic: treat the URI as a web address if it starts with http/https,
            // otherwise assume it is a local file path.
            byte[] imageData;
            if (uri.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                uri.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
            {
                using (HttpClient client = new HttpClient())
                {
                    imageData = client.GetByteArrayAsync(uri).GetAwaiter().GetResult();
                }
            }
            else
            {
                // Resolve relative paths against the current directory.
                string fullPath = Path.GetFullPath(uri);
                imageData = File.ReadAllBytes(fullPath);
            }

            // Provide the image bytes to Aspose.Words.
            args.SetData(imageData);
            return ResourceLoadingAction.UserProvided;
        }

        // For all other resource types fall back to the default loading behavior.
        return ResourceLoadingAction.Default;
    }
}

public class ReportGenerator
{
    public void Generate()
    {
        // Create a new document and attach the custom resource loading callback.
        Document doc = new Document();
        doc.ResourceLoadingCallback = new ImageUriHandler();

        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert images using URIs – the callback will convert them to Image objects internally.
        builder.InsertImage("https://example.com/images/logo.png");
        builder.InsertImage(@"C:\Images\localPicture.jpg");

        // At this point the document contains the images as Aspose.Words Shape nodes.
        // The document can now be passed to any LINQ reporting engine without further conversion.
        doc.Save("ReportWithImages.docx");
    }
}

// Example entry point.
class Program
{
    static void Main()
    {
        new ReportGenerator().Generate();
    }
}
