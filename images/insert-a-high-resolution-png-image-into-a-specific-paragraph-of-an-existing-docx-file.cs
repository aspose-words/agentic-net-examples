using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing; // Aspose.Drawing.Common namespace
using Aspose.Drawing.Imaging; // For image format if needed
using Newtonsoft.Json; // Included as required package (not used in this example)

public class Program
{
    public static void Main()
    {
        // Define file names
        const string imagePath = "sample.png";
        const string sourceDocPath = "sample.docx";
        const string resultDocPath = "result.docx";

        // -------------------------------------------------
        // 1. Create a high‑resolution PNG image (2000x2000)
        // -------------------------------------------------
        const int imgWidth = 2000;
        const int imgHeight = 2000;

        // Create bitmap and graphics objects from Aspose.Drawing
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            // Fill background with white
            graphics.Clear(Color.White);

            // (Optional) Draw a simple rectangle to make the image recognizable
            graphics.DrawRectangle(new Pen(Color.Black, 5), 100, 100, imgWidth - 200, imgHeight - 200);

            // Save the bitmap as a PNG file
            bitmap.Save(imagePath, ImageFormat.Png);
        }

        // -------------------------------------------------
        // 2. Create a sample DOCX file with several paragraphs
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Paragraph 1: Introduction");
        builder.Writeln("Paragraph 2: Target paragraph for image insertion");
        builder.Writeln("Paragraph 3: Conclusion");

        // Save the source document
        doc.Save(sourceDocPath);

        // -------------------------------------------------
        // 3. Load the existing document
        // -------------------------------------------------
        Document loadedDoc = new Document(sourceDocPath);
        DocumentBuilder docBuilder = new DocumentBuilder(loadedDoc);

        // -------------------------------------------------
        // 4. Locate the specific paragraph (second paragraph, index 1)
        // -------------------------------------------------
        Paragraph targetParagraph = (Paragraph)loadedDoc.GetChild(NodeType.Paragraph, 1, true);
        if (targetParagraph == null)
            throw new InvalidOperationException("Target paragraph not found.");

        // Move the builder cursor to the target paragraph
        docBuilder.MoveTo(targetParagraph);

        // -------------------------------------------------
        // 5. Insert the high‑resolution PNG image into the paragraph
        // -------------------------------------------------
        docBuilder.InsertImage(imagePath);

        // -------------------------------------------------
        // 6. Save the modified document
        // -------------------------------------------------
        loadedDoc.Save(resultDocPath);

        // -------------------------------------------------
        // 7. Validate that the output file was created
        // -------------------------------------------------
        if (!File.Exists(resultDocPath))
            throw new FileNotFoundException("The result document was not created.", resultDocPath);

        // (Optional) Clean up temporary files if desired
        // File.Delete(imagePath);
        // File.Delete(sourceDocPath);
    }
}
