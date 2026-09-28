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
        // Define file names.
        const string imagePath = "input.png";
        const string sourceDocPath = "input.docx";
        const string outputDocPath = "output.docx";

        // -------------------------------------------------
        // 1. Create a high‑resolution PNG image.
        // -------------------------------------------------
        const int imgWidth = 2000;
        const int imgHeight = 2000;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background with white.
                graphics.Clear(Color.White);
                // Draw a simple black rectangle for visual reference.
                graphics.DrawRectangle(Pens.Black, 100, 100, imgWidth - 200, imgHeight - 200);
            }
            // Save the image to a deterministic file.
            bitmap.Save(imagePath, ImageFormat.Png);
        }

        // Ensure the image file exists before proceeding.
        if (!File.Exists(imagePath))
            throw new FileNotFoundException("Failed to create the sample image.", imagePath);

        // -------------------------------------------------
        // 2. Create a sample DOCX file with multiple paragraphs.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Paragraph 1");
        builder.Writeln("Target Paragraph"); // This is the paragraph where we will insert the image.
        builder.Writeln("Paragraph 3");

        doc.Save(sourceDocPath);

        // Ensure the source document exists.
        if (!File.Exists(sourceDocPath))
            throw new FileNotFoundException("Failed to create the sample document.", sourceDocPath);

        // -------------------------------------------------
        // 3. Load the existing document.
        // -------------------------------------------------
        Document loadedDoc = new Document(sourceDocPath);
        DocumentBuilder docBuilder = new DocumentBuilder(loadedDoc);

        // -------------------------------------------------
        // 4. Locate the specific paragraph ("Target Paragraph").
        // -------------------------------------------------
        Paragraph targetParagraph = null;
        NodeCollection paragraphs = loadedDoc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in paragraphs)
        {
            if (para.GetText().Trim().Equals("Target Paragraph", StringComparison.Ordinal))
            {
                targetParagraph = para;
                break;
            }
        }

        if (targetParagraph == null)
            throw new InvalidOperationException("Target paragraph not found in the document.");

        // -------------------------------------------------
        // 5. Insert the high‑resolution PNG image into the target paragraph.
        // -------------------------------------------------
        docBuilder.MoveTo(targetParagraph);
        Shape insertedShape = docBuilder.InsertImage(imagePath);

        // Validate that the shape indeed contains an image.
        if (!insertedShape.HasImage)
            throw new InvalidOperationException("The inserted shape does not contain an image.");

        // -------------------------------------------------
        // 6. Save the modified document.
        // -------------------------------------------------
        loadedDoc.Save(outputDocPath);

        // -------------------------------------------------
        // 7. Validate output.
        // -------------------------------------------------
        if (!File.Exists(outputDocPath))
            throw new FileNotFoundException("The output document was not created.", outputDocPath);
    }
}
