using System;
using System.IO;
using System.Net.Http;
using Aspose.Words;
using Aspose.Words.Loading;

public class ImageUriHandler : IResourceLoadingCallback
{
    public ResourceLoadingAction ResourceLoading(ResourceLoadingArgs args)
    {
        if (args.ResourceType == ResourceType.Image)
        {
            string uri = args.OriginalUri;
            byte[] imageData;

            if (uri.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                uri.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    using var client = new HttpClient();
                    imageData = client.GetByteArrayAsync(uri).GetAwaiter().GetResult();
                }
                catch
                {
                    // If the download fails, skip loading this image.
                    return ResourceLoadingAction.Skip;
                }
            }
            else
            {
                string fullPath = Path.GetFullPath(uri);
                imageData = File.ReadAllBytes(fullPath);
            }

            args.SetData(imageData);
            return ResourceLoadingAction.UserProvided;
        }

        return ResourceLoadingAction.Default;
    }
}

public class ReportGenerator
{
    public void Generate()
    {
        // Prepare a tiny PNG image in the temp folder to guarantee a local file exists.
        string tempImagePath = Path.Combine(Path.GetTempPath(), "tempImage.png");
        if (!File.Exists(tempImagePath))
        {
            // 1x1 pixel transparent PNG.
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XcZcAAAAASUVORK5CYII=";
            File.WriteAllBytes(tempImagePath, Convert.FromBase64String(base64Png));
        }

        Document doc = new Document();
        doc.ResourceLoadingCallback = new ImageUriHandler();

        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the guaranteed local image.
        builder.InsertImage(tempImagePath);

        doc.Save("ReportWithImages.docx");
    }
}

class Program
{
    static void Main()
    {
        new ReportGenerator().Generate();
        Console.WriteLine("Report generated successfully.");
    }
}
