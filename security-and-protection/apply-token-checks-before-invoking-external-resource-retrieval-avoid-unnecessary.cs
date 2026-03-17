using System;
using System.IO;
using System.Net.Http;
using Aspose.Words;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Configure load options to use our custom resource loading callback.
        HtmlLoadOptions loadOptions = new HtmlLoadOptions
        {
            ResourceLoadingCallback = new TokenCheckingHandler()
        };

        // Sample HTML containing two images: one from an allowed host, one from a blocked host.
        string html = @"
            <html>
                <body>
                    <img src=""https://allowed.com/image1.png"" alt=""Allowed Image"" />
                    <img src=""https://blocked.com/image2.png"" alt=""Blocked Image"" />
                </body>
            </html>";

        // Load the HTML into a Document using the configured options.
        using (MemoryStream stream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(html)))
        {
            Document doc = new Document(stream, loadOptions);

            // Save the resulting document; the blocked image will be skipped.
            doc.Save("Output.docx");
        }
    }
}

// Implements IResourceLoadingCallback to perform token/host validation before any network request.
class TokenCheckingHandler : IResourceLoadingCallback
{
    // Define which hosts are considered safe for downloading resources.
    private static readonly string[] AllowedHosts = { "allowed.com", "trusted.org" };

    public ResourceLoadingAction ResourceLoading(ResourceLoadingArgs args)
    {
        // Only intervene for image resources; other types use the default loading behavior.
        if (args.ResourceType != ResourceType.Image)
            return ResourceLoadingAction.Default;

        // Verify the host of the requested URI against the whitelist.
        if (!IsHostAllowed(args.OriginalUri))
        {
            // Skip loading this resource; Aspose.Words will insert a placeholder.
            return ResourceLoadingAction.Skip;
        }

        // Host is allowed – download the image data and provide it to Aspose.Words.
        using (HttpClient client = new HttpClient())
        {
            byte[] imageData = client.GetByteArrayAsync(args.OriginalUri).GetAwaiter().GetResult();
            args.SetData(imageData);
        }

        // Indicate that we supplied the resource data ourselves.
        return ResourceLoadingAction.UserProvided;
    }

    // Helper method to check whether the URI's host is in the allowed list.
    private bool IsHostAllowed(string uri)
    {
        try
        {
            string host = new Uri(uri).Host;
            foreach (string allowed in AllowedHosts)
            {
                if (string.Equals(host, allowed, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
        }
        catch
        {
            // If the URI cannot be parsed, treat it as not allowed.
        }
        return false;
    }
}
