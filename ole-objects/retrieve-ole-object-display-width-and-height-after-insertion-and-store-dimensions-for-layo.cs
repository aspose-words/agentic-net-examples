using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare some dummy data to embed as an OLE package.
        byte[] oleData = System.Text.Encoding.UTF8.GetBytes("Hello, this is OLE content.");
        using (MemoryStream oleStream = new MemoryStream(oleData))
        {
            // Insert the OLE object into the document.
            // Parameters: stream, progId ("Package" for generic OLE package), asIcon = false, presentation = null.
            Shape oleShape = builder.InsertOleObject(oleStream, "Package", false, null);

            // After insertion, the shape's Width and Height represent the display size of the OLE object.
            double displayWidth = oleShape.Width;   // Width in points (1 point = 1/72 inch)
            double displayHeight = oleShape.Height; // Height in points

            // Store dimensions for later layout calculations (example: add a paragraph with spacing based on size).
            // Here we simply write the dimensions into the document.
            builder.Writeln($"OLE object display size: {displayWidth:F2} pt (width) x {displayHeight:F2} pt (height)");

            // Example of using the dimensions: insert a blank paragraph with a line spacing equal to the OLE height.
            Paragraph para = builder.InsertParagraph();
            para.ParagraphFormat.LineSpacing = displayHeight;
        }

        // Save the document to the local file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OleObjectDemo.docx");
        doc.Save(outputPath);
    }
}
