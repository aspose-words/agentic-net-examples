using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Drawing;

public class Program
{
    // Expected token for authorized resource loading.
    private const string ExpectedToken = "valid-token";

    public static void Main()
    {
        // Simulated token (hard‑coded for this demo).
        string suppliedToken = "valid-token"; // Change to any other value to test the token check.

        // Create a blank document.
        Document doc = new Document();

        // Attach a custom resource loading callback that validates the token before fetching external resources.
        doc.ResourceLoadingCallback = new SecureResourceLoader(suppliedToken, ExpectedToken);

        // Insert an image using a placeholder URI. The callback will intercept this request.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage("SecureImage");

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");
        doc.Save(outputPath);

        // Reload the document to verify whether the image was actually loaded.
        Document loadedDoc = new Document(outputPath);
        int imageCount = loadedDoc.GetChildNodes(NodeType.Shape, true).Count;

        // Validation: if the token was correct, the image should be present; otherwise, it should be absent.
        if (suppliedToken == ExpectedToken)
        {
            if (imageCount == 0)
                throw new InvalidOperationException("Token was valid but the image was not loaded.");
        }
        else
        {
            if (imageCount != 0)
                throw new InvalidOperationException("Invalid token allowed image loading.");
        }

        // Successful execution – no console output required.
    }
}

// Custom callback that checks a token before allowing an external image to be loaded.
public class SecureResourceLoader : IResourceLoadingCallback
{
    private readonly string _providedToken;
    private readonly string _expectedToken;

    // A tiny 1x1 PNG image (transparent) encoded in Base64.
    private static readonly byte[] PlaceholderImage = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAusB9YVh" +
        "6V8AAAAASUVORK5CYII=");

    public SecureResourceLoader(string providedToken, string expectedToken)
    {
        _providedToken = providedToken;
        _expectedToken = expectedToken;
    }

    public ResourceLoadingAction ResourceLoading(ResourceLoadingArgs args)
    {
        // Intercept only the specific placeholder used in this example.
        if (args.ResourceType == ResourceType.Image && args.OriginalUri == "SecureImage")
        {
            // Validate the token.
            if (_providedToken != _expectedToken)
                return ResourceLoadingAction.Skip; // Do not load the image.

            // Token is valid – provide the embedded placeholder image data.
            args.SetData(PlaceholderImage);
            return ResourceLoadingAction.UserProvided;
        }

        // For all other resources, use the default loading behavior.
        return ResourceLoadingAction.Default;
    }
}
