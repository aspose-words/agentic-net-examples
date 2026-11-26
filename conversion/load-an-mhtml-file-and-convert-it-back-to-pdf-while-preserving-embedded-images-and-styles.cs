using System;
using System.IO;
using Aspose.Words;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Create a sample PNG image using Aspose.Drawing.
        string imagePath = "sample.png";
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Blue);
            }
            bitmap.Save(imagePath, ImageFormat.Png);
        }

        // Build a Word document containing styled text and the image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Size = 16;
        builder.Writeln("Sample styled text with an embedded image:");
        builder.InsertImage(imagePath);

        // Save the document as MHTML.
        string mhtmlPath = "sample.mhtml";
        doc.Save(mhtmlPath, SaveFormat.Mhtml);

        // Load the MHTML file.
        Document mhtmlDoc = new Document(mhtmlPath);

        // Convert the loaded document to PDF.
        string pdfPath = "output.pdf";
        mhtmlDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
        {
            throw new InvalidOperationException("Expected output PDF was not created.");
        }
    }
}
