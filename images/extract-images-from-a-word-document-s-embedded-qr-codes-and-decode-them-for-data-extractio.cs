using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Define deterministic file names.
        const string qrImagePath = "qr.png";
        const string docPath = "sample.docx";
        const string outputFolder = "output";

        // Ensure output folder exists.
        Directory.CreateDirectory(outputFolder);

        // -------------------------------------------------
        // 1. Create a deterministic QR‑code‑like image.
        // -------------------------------------------------
        const int imgSize = 200;
        using (Bitmap bitmap = new Bitmap(imgSize, imgSize))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                // White background.
                g.Clear(Color.White);

                // Simple black square pattern to simulate a QR code.
                int blockSize = 20;
                for (int y = 0; y < imgSize; y += blockSize * 2)
                {
                    for (int x = 0; x < imgSize; x += blockSize * 2)
                    {
                        g.FillRectangle(Brushes.Black, x, y, blockSize, blockSize);
                    }
                }
            }

            // Save the image to a deterministic file.
            bitmap.Save(qrImagePath, ImageFormat.Png);
        }

        // -------------------------------------------------
        // 2. Create a Word document and insert the image.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(qrImagePath);
        doc.Save(docPath);

        // -------------------------------------------------
        // 3. Load the document and extract all images.
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        var extractedImages = new List<string>();
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Determine file extension based on image type.
            string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
            string extractedPath = Path.Combine(outputFolder, $"extracted_{imageIndex}{extension}");

            // Save the image.
            shape.ImageData.Save(extractedPath);
            extractedImages.Add(extractedPath);
            imageIndex++;
        }

        // Validate that at least one image was extracted.
        if (extractedImages.Count == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // -------------------------------------------------
        // 4. Decode each extracted image (placeholder logic).
        // -------------------------------------------------
        var decodedResults = new Dictionary<string, string>();
        foreach (string imageFile in extractedImages)
        {
            // Placeholder decode: read the first byte of the file and convert to a string.
            // In a real scenario, a QR‑code library would be used here.
            byte[] bytes = File.ReadAllBytes(imageFile);
            string decoded = $"DecodedData_{bytes.Length}";
            decodedResults[Path.GetFileName(imageFile)] = decoded;
        }

        // -------------------------------------------------
        // 5. Output the decoded data as JSON.
        // -------------------------------------------------
        string json = JsonConvert.SerializeObject(decodedResults, Formatting.Indented);
        string jsonPath = Path.Combine(outputFolder, "decoded_results.json");
        File.WriteAllText(jsonPath, json);

        // Write results to console (no interactive prompts).
        Console.WriteLine(json);
    }
}
