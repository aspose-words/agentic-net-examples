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
        // Create a sample document with text.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Font.Name = "Arial";
        builder.Font.Size = 14;
        builder.Writeln("This is a sample document with an embedded image.");

        // Create a simple bitmap using Aspose.Drawing.
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Aspose.Drawing.Color.White);
                using (Pen pen = new Pen(Aspose.Drawing.Color.Blue, 3))
                {
                    graphics.DrawEllipse(pen, 10, 10, 80, 80);
                }
            }

            // Insert the image into the document.
            using (MemoryStream imageStream = new MemoryStream())
            {
                bitmap.Save(imageStream, ImageFormat.Png);
                imageStream.Position = 0;
                builder.InsertImage(imageStream);
            }
        }

        // Save the document as MHTML.
        string mhtmlPath = "sample.mhtml";
        sourceDoc.Save(mhtmlPath, SaveFormat.Mhtml);

        // Load the MHTML file and convert it to PDF.
        Document mhtmlDoc = new Document(mhtmlPath);
        string pdfPath = "output.pdf";
        mhtmlDoc.Save(pdfPath, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(pdfPath))
        {
            throw new InvalidOperationException("Expected output PDF was not created.");
        }
    }
}
