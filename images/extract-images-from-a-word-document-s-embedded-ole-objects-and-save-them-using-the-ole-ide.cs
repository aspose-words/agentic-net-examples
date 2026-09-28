using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a deterministic sample image (sample.png)
        // -----------------------------------------------------------------
        const string imagePath = "sample.png";
        using (var bitmap = new Aspose.Drawing.Bitmap(200, 200))
        {
            using (var graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                graphics.Clear(Aspose.Drawing.Color.LightBlue);
                using (var pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.DarkBlue))
                {
                    graphics.DrawRectangle(pen, 20, 20, 160, 160);
                }
            }
            bitmap.Save(imagePath);
        }

        // -----------------------------------------------------------------
        // 2. Create a Word document and embed the image as a shape
        // -----------------------------------------------------------------
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Document with an embedded image:");
        builder.InsertImage(imagePath);

        const string docPath = "sample.docx";
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Reload the document (demonstrates extraction on a saved file)
        // -----------------------------------------------------------------
        var loadedDoc = new Document(docPath);

        // -----------------------------------------------------------------
        // 4. Extract image data from Shape nodes and save them
        // -----------------------------------------------------------------
        var shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        for (int i = 0; i < shapeNodes.Count; i++)
        {
            if (shapeNodes[i] is Shape shape && shape.HasImage)
            {
                // Determine a deterministic file name.
                string identifier = !string.IsNullOrEmpty(shape.Name) ? shape.Name : $"shape_{i}";
                string outputFileName = $"{identifier}.png";

                // Save the image data.
                shape.ImageData.Save(outputFileName);
                extractedCount++;
            }
        }

        // -----------------------------------------------------------------
        // 5. Validate that at least one image was extracted
        // -----------------------------------------------------------------
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // -----------------------------------------------------------------
        // 6. (Optional) Clean up temporary files
        // -----------------------------------------------------------------
        // File.Delete(imagePath);
        // File.Delete(docPath);
    }
}
