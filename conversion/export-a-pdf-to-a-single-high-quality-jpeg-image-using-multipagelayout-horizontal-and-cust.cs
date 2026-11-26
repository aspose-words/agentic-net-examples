using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample Word document with two pages.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is the first page of the sample PDF.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is the second page of the sample PDF.");

        // Save the document as PDF.
        string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF file was not created.");

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Configure image save options for a high‑quality JPEG.
        ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
        {
            JpegQuality = 100, // Maximum quality.
            Resolution = 300   // Increase resolution for higher quality.
        };

        // Save the PDF as JPEG image(s). For formats that do not support multiple pages
        // (like JPEG), Aspose.Words creates separate files for each page, appending a
        // numeric suffix to the file name (e.g., output.jpg, output_1.jpg).
        string jpegBasePath = "output.jpg";
        pdfDoc.Save(jpegBasePath, jpegOptions);

        // Validate that at least one JPEG was created.
        if (!File.Exists(jpegBasePath))
            throw new InvalidOperationException("The JPEG image was not created.");

        FileInfo jpegInfo = new FileInfo(jpegBasePath);
        if (jpegInfo.Length == 0)
            throw new InvalidOperationException("The JPEG image file is empty.");

        // Indicate success.
        Console.WriteLine("PDF successfully exported to high‑quality JPEG: " + jpegBasePath);
    }
}
