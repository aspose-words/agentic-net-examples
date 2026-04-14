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
        // Define file names used in the example.
        const string templatePath = "template.docx";
        const string imagePath = "apiImage.png";
        const string outputPdfPath = "result.pdf";

        // -------------------------------------------------
        // 1. Create a sample image that will simulate data
        //    received from a REST API.
        // -------------------------------------------------
        const int imgWidth = 200;
        const int imgHeight = 200;

        // Create a bitmap, clear it to white and save it.
        Bitmap bitmap = new Bitmap(imgWidth, imgHeight);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.White);
        // (Optional) draw a simple rectangle to make the image visible.
        graphics.DrawRectangle(new Pen(Aspose.Drawing.Color.Blue, 5), 10, 10, imgWidth - 20, imgHeight - 20);
        bitmap.Save(imagePath, ImageFormat.Png);
        graphics.Dispose();
        bitmap.Dispose();

        // -------------------------------------------------
        // 2. Create a simple DOCX template programmatically.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);
        templateBuilder.Writeln("=== DOCX TEMPLATE ===");
        templateBuilder.Writeln("The following image will be inserted programmatically:");
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 3. Load the template, insert the image, and save as PDF.
        // -------------------------------------------------
        Document doc = new Document(templatePath);
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the end of the document to insert the image.
        builder.MoveToDocumentEnd();
        builder.Writeln(); // Add a blank line before the image.
        builder.Writeln("Inserted image (simulated API data):");

        // Insert the image from the file created above.
        using (FileStream imgStream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
        {
            // Ensure the stream is at the beginning before insertion.
            imgStream.Position = 0;
            builder.InsertImage(imgStream);
        }

        // Save the final document as PDF.
        doc.Save(outputPdfPath, SaveFormat.Pdf);

        // -------------------------------------------------
        // 4. Validate that the PDF was created successfully.
        // -------------------------------------------------
        if (!File.Exists(outputPdfPath))
            throw new InvalidOperationException("The PDF output file was not created.");

        // (Optional) Clean up intermediate files.
        // File.Delete(templatePath);
        // File.Delete(imagePath);
    }
}
