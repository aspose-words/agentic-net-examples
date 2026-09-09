using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // File paths.
        string imagePath = Path.Combine(artifactsDir, "sample.png");
        string docPath = Path.Combine(artifactsDir, "sample.docx");
        string htmlPath = Path.Combine(artifactsDir, "output.html");

        // 1. Create a deterministic sample image.
        CreateSampleImage(imagePath);

        // 2. Create a DOCX document and insert the image.
        CreateDocumentWithImage(docPath, imagePath);

        // 3. Load the DOCX and save it as HTML with images embedded as Base64.
        SaveDocumentAsHtmlWithBase64(docPath, htmlPath);

        // Validate that the HTML file was created.
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("HTML output was not created.");
    }

    private static void CreateSampleImage(string filePath)
    {
        const int width = 200;
        const int height = 100;

        // Use Aspose.Drawing to generate the image.
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background with white.
                graphics.Clear(Color.White);
            }

            // Save the bitmap to a deterministic file name.
            bitmap.Save(filePath);
        }
    }

    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        // Create a blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the previously created image.
        builder.InsertImage(imagePath);

        // Save the document.
        doc.Save(docPath);
    }

    private static void SaveDocumentAsHtmlWithBase64(string docPath, string htmlPath)
    {
        // Load the document that contains the image.
        Document doc = new Document(docPath);

        // Configure HTML save options to embed images as Base64.
        HtmlSaveOptions options = new HtmlSaveOptions
        {
            ExportImagesAsBase64 = true,
            PrettyFormat = true
        };

        // Save as HTML; images will be embedded directly in the <img> tags.
        doc.Save(htmlPath, options);
    }
}
