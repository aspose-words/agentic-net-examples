using System;
using System.IO;
using Aspose.Words;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Paths for the sample files
        const string imagePath = "input.png";
        const string sourceDocPath = "input.docx";
        const string outputDocPath = "output.docx";

        // 1. Create a high‑resolution PNG image (2000x2000) using Aspose.Drawing
        using (Bitmap bitmap = new Bitmap(2000, 2000))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill the background with white
                graphics.Clear(Aspose.Drawing.Color.White);
                // Draw a simple rectangle to have visible content
                graphics.DrawRectangle(new Pen(Aspose.Drawing.Color.Black), 100, 100, 1800, 1800);
            }
            // Save the image to a local file
            bitmap.Save(imagePath);
        }

        // Ensure the image file was created
        if (!File.Exists(imagePath))
            throw new FileNotFoundException("Failed to create the sample image.", imagePath);

        // 2. Create a sample DOCX file with several paragraphs
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("First paragraph.");
        builder.Writeln("Target paragraph where the image will be inserted.");
        builder.Writeln("Third paragraph.");
        sourceDoc.Save(sourceDocPath);
        // No need to dispose DocumentBuilder; it does not implement IDisposable

        // Ensure the source document was created
        if (!File.Exists(sourceDocPath))
            throw new FileNotFoundException("Failed to create the source document.", sourceDocPath);

        // 3. Load the existing DOCX file
        Document doc = new Document(sourceDocPath);
        DocumentBuilder docBuilder = new DocumentBuilder(doc);

        // 4. Locate the specific paragraph (second paragraph) and insert the image
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        if (paragraphs.Count < 2)
            throw new InvalidOperationException("The document does not contain enough paragraphs.");

        Paragraph targetParagraph = (Paragraph)paragraphs[1]; // zero‑based index
        docBuilder.MoveTo(targetParagraph);
        docBuilder.InsertImage(imagePath);

        // 5. Save the modified document
        doc.Save(outputDocPath);
        // No need to dispose DocumentBuilder

        // 6. Validate that the output file exists
        if (!File.Exists(outputDocPath))
            throw new FileNotFoundException("The output document was not created.", outputDocPath);
    }
}
