using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define output directory and ensure it exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the intermediate PDF and final XPS files.
        string pdfPath = Path.Combine(outputDir, "sample.pdf");
        string xpsPath = Path.Combine(outputDir, "sample.xps");

        // -----------------------------------------------------------------
        // 1. Create a simple Word document and save it as PDF.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample PDF document created for conversion to XPS.");
        doc.Save(pdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 2. Load the generated PDF.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        // -----------------------------------------------------------------
        // 3. Convert the PDF to XPS using XpsSaveOptions.
        // -----------------------------------------------------------------
        XpsSaveOptions xpsOptions = new XpsSaveOptions();
        pdfDoc.Save(xpsPath, xpsOptions);

        // -----------------------------------------------------------------
        // 4. Validate that the XPS file was created and is not empty.
        // -----------------------------------------------------------------
        if (!File.Exists(xpsPath) || new FileInfo(xpsPath).Length == 0)
        {
            throw new InvalidOperationException("XPS conversion failed: output file is missing or empty.");
        }

        Console.WriteLine($"PDF successfully converted to XPS: {xpsPath}");
    }
}
