using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with a tiny PNG image.
        string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X6ZcAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        using (var imageStream = new MemoryStream(pngBytes))
        {
            var doc = new Document();
            var builder = new DocumentBuilder(doc);
            builder.InsertImage(imageStream);
            doc.Save("Sample.docx");
        }

        // Load the document we just created.
        var loadedDoc = new Document("Sample.docx");

        // Create a cancellation token that is already cancelled.
        var cts = new CancellationTokenSource();
        cts.Cancel();

        try
        {
            // Attempt to extract images; should stop early due to cancellation.
            ExtractImages(loadedDoc, cts.Token);
            Console.WriteLine("Image extraction completed without cancellation (unexpected).");
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Image extraction was cancelled as requested.");
        }
    }

    // Custom image extraction routine that respects cancellation.
    private static void ExtractImages(Document doc, CancellationToken token)
    {
        var shapes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapes)
        {
            // Throw if cancellation has been requested.
            token.ThrowIfCancellationRequested();

            if (shape.HasImage)
            {
                string fileName = $"ExtractedImage_{imageIndex}.png";
                shape.ImageData.Save(fileName);
                imageIndex++;
            }
        }
    }
}
