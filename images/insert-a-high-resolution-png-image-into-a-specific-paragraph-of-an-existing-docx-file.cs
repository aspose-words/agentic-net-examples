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
        const string imagePath = "input.png";
        const string sourceDocPath = "input.docx";
        const string outputDocPath = "output.docx";

        // 1. Create a high‑resolution PNG image (2000x2000) deterministically.
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
            // Save the image to a local file.
            bitmap.Save(imagePath, ImageFormat.Png);
        }

        // Ensure the image file exists before proceeding.
        if (!File.Exists(imagePath))
            throw new FileNotFoundException($"Image file was not created: {imagePath}");

        // 2. Create a sample DOCX file with three paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Paragraph 1");
        builder.Writeln("Paragraph 2"); // Target paragraph for image insertion.
        builder.Writeln("Paragraph 3");
        doc.Save(sourceDocPath);

        // Ensure the source document exists.
        if (!File.Exists(sourceDocPath))
            throw new FileNotFoundException($"Source document was not created: {sourceDocPath}");

        // 3. Load the existing document.
        Document loadedDoc = new Document(sourceDocPath);

        // 4. Locate the specific paragraph (second paragraph, index 1).
        NodeCollection paragraphs = loadedDoc.GetChildNodes(NodeType.Paragraph, true);
        if (paragraphs.Count < 2)
            throw new InvalidOperationException("The document does not contain enough paragraphs.");

        Paragraph targetParagraph = (Paragraph)paragraphs[1];

        // 5. Insert the PNG image into the target paragraph.
        DocumentBuilder insertBuilder = new DocumentBuilder(loadedDoc);
        insertBuilder.MoveTo(targetParagraph);
        insertBuilder.InsertImage(imagePath);

        // 6. Save the modified document.
        loadedDoc.Save(outputDocPath);

        // 7. Validate that the output file was created.
        if (!File.Exists(outputDocPath))
            throw new InvalidOperationException($"Failed to create output document: {outputDocPath}");

        // Cleanup: optional removal of temporary files (commented out to keep files for inspection).
        // File.Delete(imagePath);
        // File.Delete(sourceDocPath);
    }
}
