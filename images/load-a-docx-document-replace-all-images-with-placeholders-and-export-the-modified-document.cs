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
        // Define deterministic file names.
        const string inputImagePath = "input.png";
        const string placeholderImagePath = "placeholder.png";
        const string originalDocPath = "original.docx";
        const string modifiedDocPath = "modified.docx";

        // -------------------------------------------------
        // Step 1: Create a sample input image (white square).
        // -------------------------------------------------
        const int imgWidth = 100;
        const int imgHeight = 100;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
            }
            bitmap.Save(inputImagePath, ImageFormat.Png);
        }

        // -------------------------------------------------
        // Step 2: Create a placeholder image (light gray square).
        // -------------------------------------------------
        using (Bitmap placeholder = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics g = Graphics.FromImage(placeholder))
            {
                g.Clear(Color.LightGray);
            }
            placeholder.Save(placeholderImagePath, ImageFormat.Png);
        }

        // -------------------------------------------------
        // Step 3: Build a sample DOCX containing the input image.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Insert the sample image twice to have multiple images.
        builder.InsertImage(inputImagePath);
        builder.Writeln(); // Add a line break.
        builder.InsertImage(inputImagePath);
        doc.Save(originalDocPath);

        // -------------------------------------------------
        // Step 4: Load the document and replace all images with the placeholder.
        // -------------------------------------------------
        Document loadedDoc = new Document(originalDocPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                // Replace the image data with the placeholder image.
                shape.ImageData.SetImage(placeholderImagePath);
            }
        }

        // -------------------------------------------------
        // Step 5: Save the modified document.
        // -------------------------------------------------
        loadedDoc.Save(modifiedDocPath);

        // -------------------------------------------------
        // Validation: Ensure the modified document was created.
        // -------------------------------------------------
        if (!File.Exists(modifiedDocPath))
        {
            throw new Exception($"Failed to create the modified document at '{modifiedDocPath}'.");
        }

        // Optional: Clean up generated files (comment out if inspection is needed).
        // File.Delete(inputImagePath);
        // File.Delete(placeholderImagePath);
        // File.Delete(originalDocPath);
    }
}
