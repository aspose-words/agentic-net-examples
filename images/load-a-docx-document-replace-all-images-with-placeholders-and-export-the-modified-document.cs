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
        // Define file names
        const string sampleImagePath = "sample.png";
        const string placeholderImagePath = "placeholder.png";
        const string inputDocPath = "input.docx";
        const string outputDocPath = "output.docx";

        // -------------------------------------------------
        // Step 1: Create a sample image to insert into the document
        // -------------------------------------------------
        const int sampleWidth = 100;
        const int sampleHeight = 100;
        using (Bitmap sampleBitmap = new Bitmap(sampleWidth, sampleHeight))
        {
            using (Graphics g = Graphics.FromImage(sampleBitmap))
            {
                g.Clear(Color.LightBlue);
            }
            sampleBitmap.Save(sampleImagePath, ImageFormat.Png);
        }

        // -------------------------------------------------
        // Step 2: Create a placeholder image that will replace existing images
        // -------------------------------------------------
        const int placeholderWidth = 100;
        const int placeholderHeight = 100;
        using (Bitmap placeholderBitmap = new Bitmap(placeholderWidth, placeholderHeight))
        {
            using (Graphics g = Graphics.FromImage(placeholderBitmap))
            {
                g.Clear(Color.Gray);
            }
            placeholderBitmap.Save(placeholderImagePath, ImageFormat.Png);
        }

        // -------------------------------------------------
        // Step 3: Build a sample DOCX containing a few images
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Insert first image
        builder.InsertImage(sampleImagePath);
        builder.Writeln();
        // Insert second image
        builder.InsertImage(sampleImagePath);
        // Save the input document
        doc.Save(inputDocPath);

        // -------------------------------------------------
        // Step 4: Load the document and replace all images with the placeholder
        // -------------------------------------------------
        Document loadedDoc = new Document(inputDocPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int replacedCount = 0;
        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                shape.ImageData.SetImage(placeholderImagePath);
                replacedCount++;
            }
        }

        // -------------------------------------------------
        // Step 5: Save the modified document
        // -------------------------------------------------
        loadedDoc.Save(outputDocPath);

        // -------------------------------------------------
        // Validation
        // -------------------------------------------------
        if (!File.Exists(outputDocPath))
            throw new Exception("The output document was not created.");

        if (replacedCount == 0)
            throw new Exception("No images were found to replace.");

        // Clean up temporary files (optional)
        // File.Delete(sampleImagePath);
        // File.Delete(placeholderImagePath);
        // File.Delete(inputDocPath);
    }
}
