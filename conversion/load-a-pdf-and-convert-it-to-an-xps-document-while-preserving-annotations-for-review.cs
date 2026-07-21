using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample PDF file (if it does not already exist)
        string pdfPath = "sample.pdf";
        if (!File.Exists(pdfPath))
        {
            // Build a simple Word document
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello PDF with annotation.");

            // Save the document as PDF
            doc.Save(pdfPath, SaveFormat.Pdf);
        }

        // Load the PDF and convert it to XPS while preserving annotations (if supported)
        Document loadedPdf = new Document(pdfPath);
        XpsSaveOptions xpsOptions = new XpsSaveOptions
        {
            // PreserveAnnotations is not available in Aspose.Words XpsSaveOptions,
            // but we keep the options object for extensibility.
        };

        string xpsPath = "output.xps";
        loadedPdf.Save(xpsPath, xpsOptions);

        // Validate the conversion result
        if (!File.Exists(xpsPath) || new FileInfo(xpsPath).Length == 0)
        {
            throw new InvalidOperationException("XPS conversion failed or produced an empty file.");
        }

        Console.WriteLine("PDF successfully converted to XPS.");
    }
}
