using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -------------------------------------------------
        // 1. Insert a barcode image and use it as a watermark.
        // -------------------------------------------------
        // Path to the barcode image file (PNG, JPG, etc.).
        string barcodePath = @"C:\Images\barcode.png";

        // Insert the image; the method returns the Shape node.
        Shape watermark = builder.InsertImage(barcodePath);

        // Configure the shape to behave as a watermark.
        watermark.WrapType = WrapType.None;                     // No text wrapping.
        watermark.BehindText = true;                            // Place behind document text.
        watermark.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        watermark.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        watermark.Rotation = -45;                               // Optional rotation for visual effect.
        watermark.Name = "BarcodeWatermark";

        // Move the cursor to a new paragraph after the watermark.
        builder.Writeln();

        // -------------------------------------------------
        // 2. Insert an OLE object (e.g., an Excel spreadsheet).
        // -------------------------------------------------
        // Path to the OLE source file.
        string oleFilePath = @"C:\Data\SampleSpreadsheet.xlsx";

        // Insert the OLE object as an embedded object (not as an icon).
        // Parameters: file name, isLinked (false = embed), asIcon (false = show content), presentation (null = default).
        builder.Writeln("Embedded Excel Spreadsheet:");
        builder.InsertOleObject(oleFilePath, false, false, null);

        // Add a paragraph break.
        builder.Writeln();

        // -------------------------------------------------
        // 3. Insert an online video (YouTube example).
        // -------------------------------------------------
        string videoUrl = "https://www.youtube.com/watch?v=dQw4w9WgXcQ";

        // Insert the video with a default size (320x180 points).
        builder.Writeln("Online Video:");
        builder.InsertOnlineVideo(videoUrl, 320, 180);

        // -------------------------------------------------
        // Save the document to a DOCX file.
        // -------------------------------------------------
        string outputPath = @"C:\Output\GeneratedDocument.docx";
        doc.Save(outputPath, SaveFormat.Docx);
    }
}
