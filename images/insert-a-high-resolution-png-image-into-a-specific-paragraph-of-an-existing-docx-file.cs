using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class InsertHighResImage
{
    public static void Main()
    {
        // Paths for temporary files
        const string imagePath = "input.png";
        const string sourceDocPath = "input.docx";
        const string outputDocPath = "output.docx";

        // -------------------------------------------------
        // 1. Create a high‑resolution PNG image (2000x2000)
        // -------------------------------------------------
        const int imgWidth = 2000;
        const int imgHeight = 2000;

        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            // Fill background with white
            graphics.Clear(Color.White);

            // (Optional) draw a simple rectangle to make the image visible
            // graphics.DrawRectangle(new Pen(Color.Black, 5), 100, 100, imgWidth - 200, imgHeight - 200);

            // Save the bitmap as PNG
            bitmap.Save(imagePath);
        }

        // -------------------------------------------------
        // 2. Create a sample DOCX file that will act as the existing document
        // -------------------------------------------------
        Document tempDoc = new Document();
        DocumentBuilder tempBuilder = new DocumentBuilder(tempDoc);

        tempBuilder.Writeln("First paragraph.");
        tempBuilder.Writeln("Paragraph that will receive the image."); // Target paragraph
        tempBuilder.Writeln("Last paragraph.");

        tempDoc.Save(sourceDocPath);

        // -------------------------------------------------
        // 3. Load the existing document
        // -------------------------------------------------
        Document doc = new Document(sourceDocPath);
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -------------------------------------------------
        // 4. Locate the specific paragraph (second paragraph)
        // -------------------------------------------------
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        if (paragraphs.Count < 2)
            throw new InvalidOperationException("The document does not contain enough paragraphs.");

        Paragraph targetParagraph = (Paragraph)paragraphs[1]; // zero‑based index
        builder.MoveTo(targetParagraph);

        // -------------------------------------------------
        // 5. Insert the high‑resolution PNG image into the paragraph
        // -------------------------------------------------
        builder.InsertImage(imagePath);

        // -------------------------------------------------
        // 6. Save the modified document
        // -------------------------------------------------
        doc.Save(outputDocPath);

        // -------------------------------------------------
        // 7. Validate that the output file was created
        // -------------------------------------------------
        if (!File.Exists(outputDocPath))
            throw new FileNotFoundException("The output document was not created.", outputDocPath);
    }
}
