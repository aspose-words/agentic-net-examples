using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Ensure output directory exists
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);

        // Step 1: Create a sample image that represents an equation
        string equationImagePath = Path.Combine(outputDir, "equation.png");
        CreateEquationImage(equationImagePath);

        // Step 2: Create a Word document and insert the equation image
        string docPath = Path.Combine(outputDir, "sample.docx");
        CreateDocumentWithEquationImage(docPath, equationImagePath);

        // Step 3: Load the document and extract all images (including the equation image)
        Document doc = new Document(docPath);
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        int extractedCount = 0;
        foreach (Shape shape in shapes)
        {
            if (shape.HasImage)
            {
                string extractedPath = Path.Combine(outputDir, $"extracted-{extractedCount + 1}.png");
                shape.ImageData.Save(extractedPath);
                extractedCount++;
            }
        }

        // Validation: ensure at least one image was extracted
        if (extractedCount == 0)
        {
            throw new InvalidOperationException("No images were extracted from the document.");
        }

        // Indicate completion (no interactive input)
        Console.WriteLine($"Extraction complete. {extractedCount} image(s) saved to '{outputDir}'.");
    }

    private static void CreateEquationImage(string filePath)
    {
        int width = 300;
        int height = 80;
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        try
        {
            Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
            try
            {
                graphics.Clear(Aspose.Drawing.Color.White);
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24, Aspose.Drawing.FontStyle.Bold))
                {
                    graphics.DrawString(
                        "y = mx + b",
                        font,
                        new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Black),
                        new Aspose.Drawing.PointF(10, 20));
                }
            }
            finally
            {
                graphics.Dispose();
            }

            bitmap.Save(filePath, Aspose.Drawing.Imaging.ImageFormat.Png);
        }
        finally
        {
            bitmap.Dispose();
        }
    }

    private static void CreateDocumentWithEquationImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document containing an equation image:");
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }
}
