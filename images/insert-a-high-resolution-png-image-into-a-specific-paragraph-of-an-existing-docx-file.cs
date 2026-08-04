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
        // Paths for temporary files
        const string imagePath = "sample.png";
        const string sourceDocPath = "sample.docx";
        const string outputDocPath = "output.docx";

        // -------------------------------------------------
        // 1. Create a high‑resolution PNG image (2000x2000)
        // -------------------------------------------------
        const int imgWidth = 2000;
        const int imgHeight = 2000;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background with white
                graphics.Clear(Aspose.Drawing.Color.White);

                // Draw a simple red rectangle for visual reference
                using (Pen pen = new Pen(Aspose.Drawing.Color.Red, 10))
                {
                    graphics.DrawRectangle(pen, 100, 100, imgWidth - 200, imgHeight - 200);
                }
            }

            // Save the image as PNG
            bitmap.Save(imagePath, ImageFormat.Png);
        }

        // -------------------------------------------------
        // 2. Create a sample DOCX file with three paragraphs
        // -------------------------------------------------
        Document createDoc = new Document();
        DocumentBuilder createBuilder = new DocumentBuilder(createDoc);
        createBuilder.Writeln("Paragraph 1");
        createBuilder.Writeln("Paragraph 2"); // Target paragraph
        createBuilder.Writeln("Paragraph 3");
        createDoc.Save(sourceDocPath);

        // -------------------------------------------------
        // 3. Load the existing DOCX file
        // -------------------------------------------------
        Document doc = new Document(sourceDocPath);
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -------------------------------------------------
        // 4. Locate the specific paragraph (second paragraph)
        // -------------------------------------------------
        Paragraph targetParagraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 1, true);
        if (targetParagraph == null)
            throw new InvalidOperationException("Target paragraph not found.");

        // Move the builder cursor to the target paragraph
        builder.MoveTo(targetParagraph);

        // -------------------------------------------------
        // 5. Insert the high‑resolution PNG image
        // -------------------------------------------------
        Shape insertedShape = builder.InsertImage(imagePath);
        if (!insertedShape.HasImage)
            throw new InvalidOperationException("Image was not inserted correctly.");

        // -------------------------------------------------
        // 6. Save the modified document
        // -------------------------------------------------
        doc.Save(outputDocPath);

        // -------------------------------------------------
        // 7. Validate that the output file exists
        // -------------------------------------------------
        if (!File.Exists(outputDocPath))
            throw new FileNotFoundException("The output document was not created.", outputDocPath);
    }
}
