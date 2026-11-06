using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Newtonsoft.Json; // Included as required package

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample BMP image.
        const string sampleBmpPath = "sample.bmp";
        using (Bitmap bmp = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.LightBlue);
                // Simple drawing to make the image non‑empty.
                g.FillRectangle(new SolidBrush(Color.DarkBlue), 50, 50, 100, 100);
            }
            bmp.Save(sampleBmpPath);
        }

        // Step 2: Create a Word document and insert the BMP image.
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleBmpPath);
        doc.Save(docPath);

        // Step 3: Load the document and extract embedded BMP images.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            // Save the extracted image.
            string extractedPath = $"extracted_{extractedCount}.bmp";
            shape.ImageData.Save(extractedPath);
            extractedCount++;

            // Step 4: Resize the extracted BMP to 640x480.
            string resizedPath = Path.GetFileNameWithoutExtension(extractedPath) + "_resized.bmp";
            using (Bitmap original = new Bitmap(extractedPath))
            {
                using (Bitmap resized = new Bitmap(640, 480))
                {
                    using (Graphics g = Graphics.FromImage(resized))
                    {
                        g.Clear(Color.White);
                        g.DrawImage(original, 0, 0, 640, 480);
                    }
                    resized.Save(resizedPath);
                }
            }

            // Validate that the resized file was created.
            if (!File.Exists(resizedPath))
                throw new Exception($"Resized image was not created: {resizedPath}");
        }

        // Final validation: ensure at least one image was processed.
        if (extractedCount == 0)
            throw new Exception("No images were extracted from the document.");

        // Cleanup: optional removal of intermediate files (commented out).
        // File.Delete(sampleBmpPath);
        // File.Delete(docPath);
        // for (int i = 0; i < extractedCount; i++)
        // {
        //     File.Delete($"extracted_{i}.bmp");
        //     File.Delete($"extracted_{i}_resized.bmp");
        // }
    }
}
