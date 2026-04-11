using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample PDF document.
        string pdfPath = Path.Combine(outputDir, "sample.pdf");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello Aspose.Words PDF to XPS conversion.");
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify the PDF was created.
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("Failed to create the PDF file.", pdfPath);

        // Load the PDF and save it as XPS.
        Document pdfDoc = new Document(pdfPath);
        string xpsPath = Path.Combine(outputDir, "sample.xps");
        pdfDoc.Save(xpsPath, SaveFormat.Xps);

        // Verify the XPS was created.
        if (!File.Exists(xpsPath))
            throw new FileNotFoundException("Failed to create the XPS file.", xpsPath);

        Console.WriteLine("PDF successfully converted to XPS.");
        Console.WriteLine($"PDF path: {pdfPath}");
        Console.WriteLine($"XPS path: {xpsPath}");
    }
}
