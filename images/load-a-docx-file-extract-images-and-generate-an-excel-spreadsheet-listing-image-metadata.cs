using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class ImageExtractionToExcel
{
    public static void Main()
    {
        // ---------- Step 1: Create deterministic sample images ----------
        const string imagePath1 = "sample1.png";
        const string imagePath2 = "sample2.png";

        CreateSampleImage(imagePath1, 100, 100, Aspose.Drawing.Color.Blue);
        CreateSampleImage(imagePath2, 150, 80, Aspose.Drawing.Color.Green);

        // ---------- Step 2: Build a DOCX containing the sample images ----------
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.InsertImage(imagePath1);
        builder.Writeln(); // separate images with a line break
        builder.InsertImage(imagePath2);
        doc.Save(docPath);

        // ---------- Step 3: Load the DOCX and extract images ----------
        Document loadedDoc = new Document(docPath);
        var shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        var extractedImages = shapeNodes
            .OfType<Shape>()
            .Where(s => s.HasImage)
            .Select((shape, index) =>
            {
                // Determine file extension based on image type
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"extracted_{index}{extension}";
                shape.ImageData.Save(imageFileName);

                // Gather metadata
                ImageSize size = shape.ImageData.ImageSize;
                return new
                {
                    FileName = imageFileName,
                    ImageType = shape.ImageData.ImageType.ToString(),
                    WidthPixels = size.WidthPixels,
                    HeightPixels = size.HeightPixels,
                    HorizontalResolution = size.HorizontalResolution,
                    VerticalResolution = size.VerticalResolution
                };
            })
            .ToList();

        // Validation: ensure at least one image was extracted
        if (extractedImages.Count == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // ---------- Step 4: Create a CSV file (Excel‑compatible) with the metadata ----------
        const string csvPath = "ImageMetadata.csv";
        using (var writer = new StreamWriter(csvPath))
        {
            // Header row
            writer.WriteLine("FileName,ImageType,WidthPixels,HeightPixels,HorizontalResolution,VerticalResolution");

            // Data rows
            foreach (var img in extractedImages)
            {
                writer.WriteLine($"{img.FileName},{img.ImageType},{img.WidthPixels},{img.HeightPixels},{img.HorizontalResolution},{img.VerticalResolution}");
            }
        }

        // ---------- Step 5: Validate output files ----------
        if (!File.Exists(csvPath))
            throw new InvalidOperationException("CSV file was not created.");

        // Optional: clean up generated sample files (comment out if inspection is needed)
        //File.Delete(imagePath1);
        //File.Delete(imagePath2);
        //File.Delete(docPath);
        //foreach (var img in extractedImages) File.Delete(img.FileName);
    }

    // Helper method to create a deterministic PNG image using Aspose.Drawing
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color fillColor)
    {
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        {
            using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                graphics.Clear(fillColor);
            }
            bitmap.Save(filePath);
        }
    }
}
