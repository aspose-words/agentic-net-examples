using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class GenerateDocxWithBarcodeWatermarkOleVideo
{
    static void Main()
    {
        // 1. Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -----------------------------------------------------------------
        // 2. Load a pre‑generated barcode image (PNG) from disk.
        //    If you prefer to generate the barcode at runtime, add the ZXing.Net
        //    NuGet package and keep the original code that uses ZXing.
        // -----------------------------------------------------------------
        const string barcodeImagePath = @"C:\Temp\barcode.png"; // <-- ensure this file exists
        if (!File.Exists(barcodeImagePath))
            throw new FileNotFoundException($"Barcode image not found: {barcodeImagePath}");

        using (FileStream barcodeStream = File.OpenRead(barcodeImagePath))
        {
            // 3. Insert the barcode image as a watermark (behind text, centered).
            Shape watermark = builder.InsertImage(barcodeStream);
            watermark.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            watermark.RelativeVerticalPosition   = RelativeVerticalPosition.Page;
            watermark.WrapType = WrapType.None;
            watermark.BehindText = true;               // Appear as a watermark.
            watermark.HorizontalAlignment = HorizontalAlignment.Center;
            watermark.VerticalAlignment   = VerticalAlignment.Center;
            // Scale the watermark to a suitable size (adjust as needed).
            watermark.Width  = ConvertUtil.MillimeterToPoint(150);
            watermark.Height = ConvertUtil.MillimeterToPoint(50);
        }

        // -----------------------------------------------------------------
        // 4. Insert an OLE object (e.g., an Excel spreadsheet) into the document.
        // -----------------------------------------------------------------
        const string excelPath = @"C:\Temp\SampleData.xlsx";
        if (File.Exists(excelPath))
        {
            builder.InsertParagraph();
            builder.Writeln("Embedded Excel Spreadsheet:");
            // Insert the Excel file as an embedded OLE object (not as an icon).
            builder.InsertOleObject(excelPath, false, false, null);
        }
        else
        {
            builder.Writeln("[Excel file not found – skipped OLE insertion]");
        }

        // -----------------------------------------------------------------
        // 5. Insert an online video (YouTube) into the document.
        // -----------------------------------------------------------------
        builder.InsertParagraph();
        builder.Writeln("Online Video:");
        // Insert a video with default size (320x180 points).
        builder.InsertOnlineVideo("https://youtu.be/g1N9ke8Prmk", 320, 180);

        // -----------------------------------------------------------------
        // 6. Save the document to a DOCX file.
        // -----------------------------------------------------------------
        const string outputPath = @"C:\Temp\BarcodeWatermarkOleVideo.docx";
        doc.Save(outputPath, SaveFormat.Docx);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
