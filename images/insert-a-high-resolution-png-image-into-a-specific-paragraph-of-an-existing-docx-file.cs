using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // ---------- Step 1: Create a high‑resolution PNG image ----------
        const string imagePath = "input.png";
        const int imgWidth = 2000;   // pixels
        const int imgHeight = 2000;  // pixels

        // Create bitmap and graphics objects from Aspose.Drawing
        Bitmap bitmap = new Bitmap(imgWidth, imgHeight);
        Graphics graphics = Graphics.FromImage(bitmap);
        // Fill the image with white background (deterministic content)
        graphics.Clear(Color.White);
        // Optionally draw a simple rectangle to make the image visible
        graphics.DrawRectangle(Pens.Black, 100, 100, imgWidth - 200, imgHeight - 200);
        // Save the image to a local file
        bitmap.Save(imagePath);
        // Clean up drawing resources
        graphics.Dispose();
        bitmap.Dispose();

        // ---------- Step 2: Create a sample DOCX file ----------
        const string sourceDocPath = "input.docx";
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("First paragraph.");
        srcBuilder.Writeln("Paragraph that will receive the image.");
        srcBuilder.Writeln("Third paragraph.");
        sourceDoc.Save(sourceDocPath);

        // ---------- Step 3: Load the existing document ----------
        Document doc = new Document(sourceDocPath);
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Locate the specific paragraph (second paragraph, index 1)
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        if (paragraphs.Count < 2)
            throw new InvalidOperationException("The document does not contain enough paragraphs.");

        Paragraph targetParagraph = (Paragraph)paragraphs[1];
        // Move the cursor to the target paragraph
        builder.MoveTo(targetParagraph);

        // Insert the PNG image at the cursor position
        // InsertImage automatically creates a Shape and appends it to the document.
        Shape imageShape = builder.InsertImage(imagePath);
        // Ensure the shape indeed has an image
        if (!imageShape.HasImage)
            throw new InvalidOperationException("Failed to insert the image into the document.");

        // ---------- Step 4: Save the modified document ----------
        const string outputDocPath = "output.docx";
        doc.Save(outputDocPath);

        // ---------- Step 5: Validate that the output file was created ----------
        if (!File.Exists(outputDocPath))
            throw new FileNotFoundException("The output document was not created.", outputDocPath);
    }
}
