using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Simulated authentication token.
        string providedToken = "valid-token";
        string requiredToken = "valid-token";

        // Create a new blank document.
        Document doc = new Document();

        // Assign a custom resource loading callback that checks the token before providing an image.
        doc.ResourceLoadingCallback = new ImageTokenHandler(providedToken, requiredToken);

        // Build the document content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document with conditional external image:");
        // The placeholder URI triggers the callback.
        builder.InsertImage("ProtectedImage");

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the output document.");

        // Load the saved document to ensure it can be opened.
        Document loaded = new Document(outputPath);
        Console.WriteLine("Document saved and loaded successfully.");
    }
}

// Custom callback that provides an image only when a valid token is supplied.
// Instead of downloading from the internet, it uses an embedded 1x1 PNG to avoid SSL issues.
public class ImageTokenHandler : IResourceLoadingCallback
{
    private readonly string _providedToken;
    private readonly string _requiredToken;

    // A tiny 1x1 pixel transparent PNG (base64 encoded).
    private const string Base64Png = 
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO7+X3cAAAAASUVORK5CYII=";

    public ImageTokenHandler(string providedToken, string requiredToken)
    {
        _providedToken = providedToken;
        _requiredToken = requiredToken;
    }

    public ResourceLoadingAction ResourceLoading(ResourceLoadingArgs args)
    {
        // Only handle image resources with the specific placeholder URI.
        if (args.ResourceType == ResourceType.Image && args.OriginalUri == "ProtectedImage")
        {
            // Check the token before providing any image data.
            if (_providedToken == _requiredToken)
            {
                // Token is valid – provide the embedded PNG data.
                byte[] imageData = Convert.FromBase64String(Base64Png);
                args.SetData(imageData);
                return ResourceLoadingAction.UserProvided;
            }
            else
            {
                // Token is invalid – skip loading the external resource.
                return ResourceLoadingAction.Skip;
            }
        }

        // For all other resources, use the default loading behavior.
        return ResourceLoadingAction.Default;
    }
}
