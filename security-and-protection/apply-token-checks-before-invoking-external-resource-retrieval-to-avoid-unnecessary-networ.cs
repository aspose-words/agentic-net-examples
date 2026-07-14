using System;
using System.IO;
using System.Net.Http;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Expected token that authorizes external resource loading.
        const string expectedToken = "valid-token";

        // Create a blank document.
        Document doc = new Document();

        // Assign a custom resource loading callback that checks the token before downloading.
        doc.ResourceLoadingCallback = new TokenCheckingCallback(expectedToken);

        // Insert an image using a URL. The callback will decide whether to load it.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage("https://www.example.com/sample-image.png");

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
        doc.Save(outputPath, SaveFormat.Docx);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the output document.");

        Console.WriteLine($"Document saved to: {outputPath}");
    }

    // Callback that validates a token before allowing image download.
    private class TokenCheckingCallback : IResourceLoadingCallback
    {
        private readonly string _validToken;

        public TokenCheckingCallback(string validToken)
        {
            _validToken = validToken;
        }

        public ResourceLoadingAction ResourceLoading(ResourceLoadingArgs args)
        {
            // Only intercept image resources.
            if (args.ResourceType == ResourceType.Image)
            {
                // Simulate token retrieval (in a real scenario, obtain it from a secure source).
                string providedToken = GetProvidedToken();

                // If the token is invalid, skip loading the external resource.
                if (!string.Equals(providedToken, _validToken, StringComparison.Ordinal))
                    return ResourceLoadingAction.Skip;

                // Token is valid – attempt to download the image.
                try
                {
                    using (HttpClient client = new HttpClient())
                    {
                        // Synchronously get the image bytes; any failure will be caught below.
                        byte[] imageData = client.GetByteArrayAsync(args.OriginalUri).GetAwaiter().GetResult();
                        args.SetData(imageData);
                    }

                    return ResourceLoadingAction.UserProvided;
                }
                catch (HttpRequestException)
                {
                    // If the request fails (e.g., 404), skip loading to avoid an exception.
                    return ResourceLoadingAction.Skip;
                }
                catch (Exception)
                {
                    // For any other unexpected errors, also skip loading.
                    return ResourceLoadingAction.Skip;
                }
            }

            // For all other resource types, use the default loading behavior.
            return ResourceLoadingAction.Default;
        }

        // Placeholder method to obtain a token. Replace with real logic as needed.
        private string GetProvidedToken()
        {
            // For demonstration, we return the same token that authorizes loading.
            return "valid-token";
        }
    }
}
