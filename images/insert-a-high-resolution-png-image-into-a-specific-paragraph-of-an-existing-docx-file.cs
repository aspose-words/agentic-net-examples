using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class InsertHighResImageExample
{
    public static void Main()
    {
        // Paths for temporary files
        const string imagePath = "highres.png";
        const string sourceDocPath = "source.docx";
        const string resultDocPath = "result.docx";

        // -------------------------------------------------
        // 1. Create a high‑resolution PNG image (e.g., 2000×2000)
        // -------------------------------------------------
        int width = 2000;
        int height = 2000;
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill with a solid color for visibility
                graphics.Clear(Aspose.Drawing.Color.LightBlue);
            }

            // Save the image to a deterministic file name
            bitmap.Save(imagePath, ImageFormat.Png);
        }

        // -------------------------------------------------
        // 2. Create a sample DOCX file with several paragraphs
        // -------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        builder.Writeln("Paragraph 1: Introduction.");
        builder.Writeln("Paragraph 2: Target location for the image.");
        builder.Writeln("Paragraph 3: Conclusion.");

        sourceDoc.Save(sourceDocPath);

        // -------------------------------------------------
        // 3. Load the existing document and insert the image
        // -------------------------------------------------
        Document doc = new Document(sourceDocPath);
        DocumentBuilder docBuilder = new DocumentBuilder(doc);

        // Locate the specific paragraph (Paragraph 2, zero‑based index = 1)
        Node targetParagraph = doc.GetChild(NodeType.Paragraph, 1, true);
        if (targetParagraph == null)
            throw new InvalidOperationException("Target paragraph not found.");

        // Move the builder cursor to the target paragraph
        docBuilder.MoveTo(targetParagraph);

        // Insert the high‑resolution PNG image inline
        Shape imageShape = docBuilder.InsertImage(imagePath);

        // Optional: adjust image size if needed (e.g., 300 points width)
        // imageShape.Width = 300;
        // imageShape.Height = 300;

        // -------------------------------------------------
        // 4. Save the modified document
        // -------------------------------------------------
        doc.Save(resultDocPath);

        // Simple validation that the output file was created
        if (!File.Exists(resultDocPath))
            throw new FileNotFoundException("Result document was not created.", resultDocPath);
    }
}
