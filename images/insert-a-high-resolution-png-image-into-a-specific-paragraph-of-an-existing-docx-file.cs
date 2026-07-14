using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class InsertImageExample
{
    public static void Main()
    {
        // Paths for the files used in the example.
        const string imagePath = "highres.png";
        const string sourceDocPath = "sample.docx";
        const string outputDocPath = "output.docx";

        // -------------------------------------------------
        // 1. Create a high‑resolution PNG image.
        // -------------------------------------------------
        const int imgWidth = 2000;
        const int imgHeight = 2000;

        // Create a bitmap and a graphics object to draw on it.
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            // Fill the background with white.
            graphics.Clear(Color.White);

            // (Optional) Draw a simple rectangle to make the image recognizable.
            // The rectangle is drawn in a light gray color.
            graphics.FillRectangle(new SolidBrush(Color.LightGray), 100, 100, imgWidth - 200, imgHeight - 200);

            // Save the bitmap as a PNG file.
            bitmap.Save(imagePath);
        }

        // Verify that the image file was created.
        if (!File.Exists(imagePath))
            throw new FileNotFoundException($"Image file '{imagePath}' was not created.");

        // -------------------------------------------------
        // 2. Create a sample DOCX file with several paragraphs.
        // -------------------------------------------------
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);

        builder.Writeln("First paragraph.");
        builder.Writeln("Target paragraph where the image will be inserted.");
        builder.Writeln("Third paragraph.");

        sampleDoc.Save(sourceDocPath);

        // Verify that the source document was created.
        if (!File.Exists(sourceDocPath))
            throw new FileNotFoundException($"Source document '{sourceDocPath}' was not created.");

        // -------------------------------------------------
        // 3. Load the existing document and insert the image into the target paragraph.
        // -------------------------------------------------
        Document doc = new Document(sourceDocPath);
        DocumentBuilder docBuilder = new DocumentBuilder(doc);

        // Move the cursor to the second paragraph (index 1, zero‑based).
        docBuilder.MoveToParagraph(1, 0);

        // Insert the high‑resolution PNG image inline.
        Shape insertedShape = docBuilder.InsertImage(imagePath);

        // Optional: verify that the shape indeed contains an image.
        if (!insertedShape.HasImage)
            throw new InvalidOperationException("The inserted shape does not contain an image.");

        // -------------------------------------------------
        // 4. Save the modified document.
        // -------------------------------------------------
        doc.Save(outputDocPath);

        // Validate that the output file exists.
        if (!File.Exists(outputDocPath))
            throw new FileNotFoundException($"Output document '{outputDocPath}' was not saved.");

        // The example finishes without requiring any user interaction.
    }
}
