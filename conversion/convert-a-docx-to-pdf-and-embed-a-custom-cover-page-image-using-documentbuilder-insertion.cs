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
        // Create a simple cover image using Aspose.Drawing.
        const string coverImagePath = "cover.png";
        using (Bitmap bitmap = new Bitmap(600, 800))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.LightBlue);
                // Additional drawing can be added here if desired.
            }
            bitmap.Save(coverImagePath, ImageFormat.Png);
        }

        // Create a sample DOCX document.
        const string inputDocxPath = "input.docx";
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);
        sourceBuilder.Writeln("This is the main document content.");
        sourceDoc.Save(inputDocxPath, SaveFormat.Docx);

        // Load the DOCX, insert the cover image, and convert to PDF.
        Document doc = new Document(inputDocxPath);
        DocumentBuilder coverBuilder = new DocumentBuilder(doc);
        coverBuilder.MoveToDocumentStart();
        coverBuilder.InsertImage(coverImagePath);
        coverBuilder.InsertBreak(BreakType.PageBreak);

        const string outputPdfPath = "output.pdf";
        doc.Save(outputPdfPath, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(outputPdfPath))
        {
            throw new InvalidOperationException("Expected output PDF was not created.");
        }
    }
}
