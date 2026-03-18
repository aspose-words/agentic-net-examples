using System;
using System.IO;
using System.Net.Http;
using Aspose.Words;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Assign the custom resource loading callback.
        doc.ResourceLoadingCallback = new WebServiceImageLoader();

        // Build the document and insert images using placeholder URIs.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage("logo");      // Will be replaced by the image returned from the web service.
        builder.InsertImage("banner");    // Another placeholder.

        // Save the resulting document.
        doc.Save("Result.docx");
    }
}

// Implements IResourceLoadingCallback to fetch images from a web service.
class WebServiceImageLoader : IResourceLoadingCallback
{
    // Base URL of the web service that returns image bytes.
    private const string ServiceBaseUrl = "https://example.com/api/images/";

    public ResourceLoadingAction ResourceLoading(ResourceLoadingArgs args)
    {
        // Process only image resources.
        if (args.ResourceType == ResourceType.Image)
        {
            // Construct the request URL using the original URI as an identifier.
            string requestUrl = ServiceBaseUrl + Uri.EscapeDataString(args.OriginalUri);

            using (HttpClient client = new HttpClient())
            {
                // Synchronously download the image bytes.
                byte[] imageBytes = client.GetByteArrayAsync(requestUrl).GetAwaiter().GetResult();

                // Provide the downloaded bytes to Aspose.Words.
                args.SetData(imageBytes);
            }

            // Instruct Aspose.Words to use the user‑provided data.
            return ResourceLoadingAction.UserProvided;
        }

        // For all other resource types, fall back to the default loading behavior.
        return ResourceLoadingAction.Default;
    }
}
