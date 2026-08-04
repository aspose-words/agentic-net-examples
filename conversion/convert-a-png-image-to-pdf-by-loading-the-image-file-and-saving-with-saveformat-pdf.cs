using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Path for the temporary PNG image.
        const string pngPath = "sample.png";

        // Base64‑encoded 1×1 red PNG image.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BFwAE/wJ/6ZcKAAAAAElFTkSuQmCC";

        // Write the PNG file to disk.
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(pngPath, pngBytes);

        // Create a new blank Word document.
        Document doc = new Document();

        // Insert the PNG image into the document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(pngPath);

        // Save the document as PDF.
        const string pdfPath = "output.pdf";
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF file was not created.");

        // Clean up temporary files (optional).
        File.Delete(pngPath);
    }
}
