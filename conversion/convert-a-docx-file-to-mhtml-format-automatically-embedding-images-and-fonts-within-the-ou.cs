using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
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
                graphics.Clear(Color.LightBlue);
                using (Pen pen = new Pen(Color.DarkBlue, 3))
                {
                    graphics.DrawEllipse(pen, 10, 10, 80, 80);
                }
            }

            bitmap.Save(imagePath, ImageFormat.Png);
        }

        // Create a sample DOCX document that includes text and the image.
        string docxPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial";
        builder.Writeln("This is a sample document with an embedded image and a custom font.");
        builder.InsertImage(imagePath);
        doc.Save(docxPath, SaveFormat.Docx);

        // Load the DOCX document for conversion.
        Document loadedDoc = new Document(docxPath);

        // Configure HTML save options to produce MHTML with embedded images and fonts.
        HtmlSaveOptions mhtmlOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            ExportImagesAsBase64 = true,
            ExportFontsAsBase64 = true
        };

        // Save the document as MHTML.
        string mhtmlPath = "output.mhtml";
        loadedDoc.Save(mhtmlPath, mhtmlOptions);

        // Verify that the MHTML file was created.
        if (!File.Exists(mhtmlPath))
        {
            throw new InvalidOperationException("Expected MHTML output was not created.");
        }
    }
}
