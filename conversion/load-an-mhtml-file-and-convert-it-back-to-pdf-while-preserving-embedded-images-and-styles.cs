using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;

public class Program
{
    public static void Main()
    {
        // File names used in the example
        const string imagePath = "sample.png";
        const string htmlPath = "sample.html";
        const string mhtmlPath = "sample.mht";
        const string pdfPath = "output.pdf";

        // -----------------------------------------------------------------
        // 1. Create a simple PNG image using Aspose.Drawing (no System.Drawing)
        // -----------------------------------------------------------------
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            // Create a Graphics object from the bitmap (Aspose.Drawing API)
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Blue);
                using (Pen pen = new Pen(Color.Yellow, 5))
                {
                    graphics.DrawRectangle(pen, 10, 10, 80, 80);
                }
            }

            // Save the bitmap as PNG
            bitmap.Save(imagePath, ImageFormat.Png);
        }

        // -----------------------------------------------------------------
        // 2. Build an HTML document that references the image and contains CSS
        // -----------------------------------------------------------------
        string htmlContent = $@"
<!DOCTYPE html>
<html>
<head>
    <style>
        .title {{ color: red; font-size: 24px; }}
    </style>
</head>
<body>
    <h1 class='title'>Sample MHTML Document</h1>
    <p>This document contains an image and styled text.</p>
    <img src='{imagePath}' alt='Sample Image' />
</body>
</html>";
        File.WriteAllText(htmlPath, htmlContent);

        // -----------------------------------------------------------------
        // 3. Load the HTML into an Aspose.Words Document
        // -----------------------------------------------------------------
        Document doc = new Document(htmlPath);

        // -----------------------------------------------------------------
        // 4. Save the document as MHTML, embedding images and styles
        // -----------------------------------------------------------------
        HtmlSaveOptions mhtmlOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            // Ensure resources are embedded in the MHTML package
            ExportCidUrlsForMhtmlResources = false,
            ExportImagesAsBase64 = false,
            ExportFontResources = false
        };
        doc.Save(mhtmlPath, mhtmlOptions);

        // -----------------------------------------------------------------
        // 5. Load the generated MHTML file
        // -----------------------------------------------------------------
        Document mhtmlDoc = new Document(mhtmlPath);

        // -----------------------------------------------------------------
        // 6. Convert the MHTML document to PDF while preserving content
        // -----------------------------------------------------------------
        mhtmlDoc.Save(pdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 7. Validate that the PDF was created successfully
        // -----------------------------------------------------------------
        if (!File.Exists(pdfPath) || new FileInfo(pdfPath).Length == 0)
        {
            throw new InvalidOperationException("PDF conversion failed: output file is missing or empty.");
        }

        Console.WriteLine("PDF conversion succeeded. Output file: " + pdfPath);
    }
}
