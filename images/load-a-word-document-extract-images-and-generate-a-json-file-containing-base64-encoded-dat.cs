using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class ImageExtractionExample
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a deterministic sample image (input.png).
        // -----------------------------------------------------------------
        string imagePath = Path.Combine(artifactsDir, "input.png");
        CreateSampleImage(imagePath, 200, 200);

        // -----------------------------------------------------------------
        // 2. Create a Word document and insert the sample image.
        // -----------------------------------------------------------------
        string docPath = Path.Combine(artifactsDir, "sample.docx");
        CreateDocumentWithImage(docPath, imagePath);

        // -----------------------------------------------------------------
        // 3. Load the document and extract all images as Base64 strings.
        // -----------------------------------------------------------------
        var extractedImages = ExtractImagesAsBase64(docPath);

        // -----------------------------------------------------------------
        // 4. Serialize the extracted data to JSON and save it.
        // -----------------------------------------------------------------
        string jsonPath = Path.Combine(artifactsDir, "images.json");
        SaveJson(extractedImages, jsonPath);

        // Validate that the JSON file was created.
        if (!File.Exists(jsonPath))
            throw new InvalidOperationException("Failed to create the JSON output file.");

        Console.WriteLine($"Extraction complete. JSON saved to: {jsonPath}");
    }

    // Creates a simple white bitmap and saves it to the specified file.
    private static void CreateSampleImage(string fileName, int width, int height)
    {
        // Ensure any previous file is removed.
        if (File.Exists(fileName))
            File.Delete(fileName);

        // Create bitmap and graphics objects using Aspose.Drawing.
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.White);

        // Save the bitmap to disk.
        bitmap.Save(fileName);

        // Clean up.
        graphics.Dispose();
        bitmap.Dispose();

        // Validate that the image file exists.
        if (!File.Exists(fileName))
            throw new InvalidOperationException($"Failed to create sample image at {fileName}");
    }

    // Creates a new Word document, inserts the image, and saves it.
    private static void CreateDocumentWithImage(string docFileName, string imageFileName)
    {
        // Ensure any previous document is removed.
        if (File.Exists(docFileName))
            File.Delete(docFileName);

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image using the builder (InsertImage handles shape creation internally).
        builder.InsertImage(imageFileName);

        // Save the document.
        doc.Save(docFileName);

        // Validate that the document file exists.
        if (!File.Exists(docFileName))
            throw new InvalidOperationException($"Failed to create Word document at {docFileName}");
    }

    // Represents a single extracted image entry for JSON serialization.
    private class ImageInfo
    {
        public string FileName { get; set; }
        public string Base64Data { get; set; }
    }

    // Loads the document, extracts images, and returns a list of ImageInfo objects.
    private static List<ImageInfo> ExtractImagesAsBase64(string docFileName)
    {
        Document doc = new Document(docFileName);

        // Get all Shape nodes in the document.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        var images = new List<ImageInfo>();
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            // Process only shapes that actually contain an image.
            if (!shape.HasImage)
                continue;

            // Determine a deterministic file name for the image.
            string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
            string imageFileName = $"image{imageIndex}{extension}";

            // Save the image to a memory stream.
            using (MemoryStream ms = new MemoryStream())
            {
                shape.ImageData.Save(ms);
                ms.Position = 0; // Reset before reading.

                byte[] imageBytes = ms.ToArray();
                string base64 = Convert.ToBase64String(imageBytes);

                images.Add(new ImageInfo
                {
                    FileName = imageFileName,
                    Base64Data = base64
                });
            }

            imageIndex++;
        }

        // Validation: at least one image must have been extracted.
        if (images.Count == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        return images;
    }

    // Serializes the list of images to JSON and writes it to a file.
    private static void SaveJson(List<ImageInfo> images, string jsonFilePath)
    {
        var options = new JsonSerializerOptions { WriteIndented = true };
        string json = JsonSerializer.Serialize(images, options);
        File.WriteAllText(jsonFilePath, json, Encoding.UTF8);
    }
}
