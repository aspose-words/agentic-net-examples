using System;
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
        builder.InsertImage("logo");      // Will be replaced by the image returned from the loader.
        builder.InsertImage("banner");    // Another placeholder.

        // Save the resulting document.
        doc.Save("Result.docx");
    }
}

// Implements IResourceLoadingCallback to provide image data.
class WebServiceImageLoader : IResourceLoadingCallback
{
    // A 1x1 pixel PNG (transparent) encoded in base64.
    private static readonly byte[] PlaceholderImage = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X9WcAAAAASUVORK5CYII=");

    public ResourceLoadingAction ResourceLoading(ResourceLoadingArgs args)
    {
        // Process only image resources.
        if (args.ResourceType == ResourceType.Image)
        {
            // Provide the placeholder image bytes to Aspose.Words.
            args.SetData(PlaceholderImage);
            return ResourceLoadingAction.UserProvided;
        }

        // For all other resource types, fall back to the default loading behavior.
        return ResourceLoadingAction.Default;
    }
}
