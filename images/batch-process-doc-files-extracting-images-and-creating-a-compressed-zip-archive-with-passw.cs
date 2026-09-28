using System;
using System.IO;
using System.IO.Compression;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class BatchImageExtractor
{
    private const string InputDocsFolder = "InputDocs";
    private const string ImagesFolder = "ExtractedImages";
    private const string ZipPath = "ImagesArchive.zip";
    private const string ZipPassword = "Secret123"; // retained for intent, not used by built‑in ZIP

    public static void Main()
    {
        // Ensure a clean environment.
        CleanupFolders();

        // Create a deterministic sample image.
        string sampleImagePath = "sample.png";
        CreateSampleImage(sampleImagePath);

        // Create sample DOCX files that contain the image.
        Directory.CreateDirectory(InputDocsFolder);
        CreateSampleDocument(Path.Combine(InputDocsFolder, "Doc1.docx"), sampleImagePath);
        CreateSampleDocument(Path.Combine(InputDocsFolder, "Doc2.docx"), sampleImagePath);

        // Extract images from all DOCX files in the input folder.
        Directory.CreateDirectory(ImagesFolder);
        int extractedCount = ExtractImagesFromDocuments(InputDocsFolder, ImagesFolder);

        // Validate that at least one image was extracted.
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the documents.");

        // Create a ZIP archive of the extracted images.
        CreateZip(ImagesFolder, ZipPath);

        // Validate ZIP creation.
        if (!File.Exists(ZipPath))
            throw new FileNotFoundException("Failed to create the ZIP archive.");

        // Optional cleanup of the temporary sample image.
        // File.Delete(sampleImagePath);
    }

    private static void CleanupFolders()
    {
        if (Directory.Exists(InputDocsFolder))
            Directory.Delete(InputDocsFolder, true);
        if (Directory.Exists(ImagesFolder))
            Directory.Delete(ImagesFolder, true);
        if (File.Exists(ZipPath))
            File.Delete(ZipPath);
        if (File.Exists("sample.png"))
            File.Delete("sample.png");
    }

    private static void CreateSampleImage(string path)
    {
        const int width = 200;
        const int height = 200;

        // Use Aspose.Drawing to create a deterministic PNG image.
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap);
        g.Clear(Aspose.Drawing.Color.White);
        using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Red, 5))
        {
            g.DrawRectangle(pen, 10, 10, width - 20, height - 20);
        }

        // Save the bitmap as PNG.
        bitmap.Save(path, Aspose.Drawing.Imaging.ImageFormat.Png);

        // Clean up drawing resources.
        g.Dispose();
        bitmap.Dispose();
    }

    private static void CreateSampleDocument(string docPath, string imagePath)
    {
        // Build a simple document that contains the sample image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample document containing an image:");
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    private static int ExtractImagesFromDocuments(string docsFolder, string outputFolder)
    {
        int imageIndex = 0;

        foreach (string filePath in Directory.GetFiles(docsFolder, "*.docx"))
        {
            Document doc = new Document(filePath);
            NodeCollection shapes = doc.GetChildNodes(Aspose.Words.NodeType.Shape, true);

            foreach (Shape shape in shapes)
            {
                if (shape.HasImage)
                {
                    string imageFileName = $"{Path.GetFileNameWithoutExtension(filePath)}_image_{imageIndex}.png";
                    string imageFullPath = Path.Combine(outputFolder, imageFileName);
                    shape.ImageData.Save(imageFullPath);
                    imageIndex++;
                }
            }
        }

        return imageIndex;
    }

    private static void CreateZip(string sourceFolder, string zipFilePath)
    {
        // Use built‑in .NET compression to create a ZIP archive.
        // Password protection is not supported by System.IO.Compression;
        // the password variable is retained to reflect the original intent.
        if (File.Exists(zipFilePath))
            File.Delete(zipFilePath);

        ZipFile.CreateFromDirectory(sourceFolder, zipFilePath, CompressionLevel.Optimal, false);
    }
}
