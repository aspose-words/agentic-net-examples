using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a new blank document.
        // -----------------------------------------------------------------
        Document doc = new Document();

        // -----------------------------------------------------------------
        // 2. Add a text watermark with custom formatting.
        // -----------------------------------------------------------------
        TextWatermarkOptions textOptions = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 48,
            Color = Color.DarkGray,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = false // make it fully opaque
        };
        doc.Watermark.SetText("CONFIDENTIAL", textOptions);

        // -----------------------------------------------------------------
        // 3. Save the document that now contains the text watermark.
        // -----------------------------------------------------------------
        string textWatermarkPath = Path.Combine(Environment.CurrentDirectory, "TextWatermark.docx");
        doc.Save(textWatermarkPath);

        // -----------------------------------------------------------------
        // 4. Load the previously saved document.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(textWatermarkPath);

        // -----------------------------------------------------------------
        // 5. Replace the text watermark with an image watermark.
        //    First, remove the existing watermark if it is a text watermark.
        // -----------------------------------------------------------------
        if (loadedDoc.Watermark.Type == WatermarkType.Text)
        {
            loadedDoc.Watermark.Remove();
        }

        // -----------------------------------------------------------------
        // 6. Define image watermark options (scale and washout).
        // -----------------------------------------------------------------
        ImageWatermarkOptions imageOptions = new ImageWatermarkOptions
        {
            Scale = 5,               // enlarge the image 5 times
            IsWashout = false        // keep original colors
        };

        // -----------------------------------------------------------------
        // 7. Add the image watermark from a file.
        // -----------------------------------------------------------------
        string imagePath = Path.Combine(Environment.CurrentDirectory, "logo.png"); // ensure this file exists
        loadedDoc.Watermark.SetImage(imagePath, imageOptions);

        // -----------------------------------------------------------------
        // 8. Save the document that now contains the image watermark.
        // -----------------------------------------------------------------
        string imageWatermarkPath = Path.Combine(Environment.CurrentDirectory, "ImageWatermark.docx");
        loadedDoc.Save(imageWatermarkPath);
    }
}
