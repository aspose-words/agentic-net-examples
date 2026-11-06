using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Replacing;
using Aspose.Drawing;
using Newtonsoft.Json;

public class ReplaceLowResolutionImages
{
    public static void Main()
    {
        // Prepare sample images.
        const string lowResPath = "low.png";
        const string highResPath = "high.png";

        CreateSampleImage(lowResPath, 100, 100, Aspose.Drawing.Color.LightGray);
        CreateSampleImage(highResPath, 500, 500, Aspose.Drawing.Color.LightBlue);

        // Create a sample Word document containing low‑resolution images.
        const string inputDocPath = "input.docx";
        CreateDocumentWithImages(inputDocPath, lowResPath, 3);

        // Load the document.
        Document doc = new Document(inputDocPath);

        // Replace low‑resolution images with high‑resolution versions.
        int replacedCount = 0;
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            if (!shape.HasImage) continue;

            // Determine image size in pixels.
            ImageSize size = shape.ImageData.ImageSize;

            // Consider images with both dimensions <= 150 as low‑resolution.
            if (size.WidthPixels <= 150 && size.HeightPixels <= 150)
            {
                shape.ImageData.SetImage(highResPath);
                replacedCount++;
            }
        }

        // Validate that at least one image was replaced.
        if (replacedCount == 0)
            throw new InvalidOperationException("No low‑resolution images were found to replace.");

        // Save the updated document.
        const string outputDocPath = "output.docx";
        doc.Save(outputDocPath);

        // Validate output file existence.
        if (!File.Exists(outputDocPath))
            throw new FileNotFoundException("The output document was not created.", outputDocPath);

        // Example usage of Newtonsoft.Json to satisfy package requirement.
        var resultInfo = new { ReplacedImages = replacedCount, OutputFile = outputDocPath };
        string json = JsonConvert.SerializeObject(resultInfo);
        File.WriteAllText("result.json", json);
    }

    private static void CreateSampleImage(string path, int width, int height, Aspose.Drawing.Color backColor)
    {
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                g.Clear(backColor);
            }
            bitmap.Save(path);
        }
    }

    private static void CreateDocumentWithImages(string docPath, string imagePath, int repeatCount)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 0; i < repeatCount; i++)
        {
            builder.Writeln($"Image #{i + 1}:");
            builder.InsertImage(imagePath);
            builder.Writeln(); // Add spacing.
        }

        doc.Save(docPath);
    }
}
