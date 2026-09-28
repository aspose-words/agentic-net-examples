using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        string imagePath = "sample.png";
        string inputDocPath = "input.docx";
        string outputHtmlPath = "output.html";

        // -------------------------------------------------
        // Create a deterministic sample image (100x100 white)
        // -------------------------------------------------
        int imgWidth = 100;
        int imgHeight = 100;
        Bitmap bitmap = new Bitmap(imgWidth, imgHeight);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.White);
        // Optionally draw something simple
        // Save the image
        bitmap.Save(imagePath);
        graphics.Dispose();
        bitmap.Dispose();

        // -------------------------------------------------
        // Create a DOCX document and insert the sample image twice
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document with images:");
        builder.InsertImage(imagePath);
        builder.Writeln();
        builder.InsertImage(imagePath);
        doc.Save(inputDocPath);

        // -------------------------------------------------
        // Load the DOCX, replace all images with placeholders
        // -------------------------------------------------
        Document loadedDoc = new Document(inputDocPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                // Create a placeholder rectangle with same size
                Shape placeholder = new Shape(loadedDoc, ShapeType.Rectangle);
                placeholder.Width = shape.Width;
                placeholder.Height = shape.Height;
                placeholder.WrapType = shape.WrapType;
                placeholder.RelativeHorizontalPosition = shape.RelativeHorizontalPosition;
                placeholder.RelativeVerticalPosition = shape.RelativeVerticalPosition;
                placeholder.HorizontalAlignment = shape.HorizontalAlignment;
                placeholder.VerticalAlignment = shape.VerticalAlignment;

                // Add text "[Image]" inside the placeholder
                Paragraph para = new Paragraph(loadedDoc);
                Run run = new Run(loadedDoc, "[Image]");
                para.AppendChild(run);
                placeholder.AppendChild(para);

                // Replace the original image shape with the placeholder
                shape.ParentNode.InsertAfter(placeholder, shape);
                shape.Remove();
            }
        }

        // -------------------------------------------------
        // Save the modified document as HTML preserving layout
        // -------------------------------------------------
        loadedDoc.Save(outputHtmlPath, SaveFormat.Html);

        // -------------------------------------------------
        // Validate that the HTML file was created
        // -------------------------------------------------
        if (!File.Exists(outputHtmlPath))
            throw new Exception("HTML output was not created.");

        // Clean up temporary files (optional)
        // File.Delete(imagePath);
        // File.Delete(inputDocPath);
    }
}
