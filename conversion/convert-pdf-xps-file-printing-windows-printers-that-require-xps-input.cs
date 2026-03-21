using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace PdfToXpsConversion
{
    class Program
    {
        static void Main(string[] args)
        {
            // Determine working directory.
            string workingDir = Directory.GetCurrentDirectory();

            // Paths for input PDF and output XPS.
            string pdfPath = Path.Combine(workingDir, "sample.pdf");
            string xpsPath = Path.Combine(workingDir, "sample.xps");

            // Ensure the input PDF exists; if not, create a simple one.
            if (!File.Exists(pdfPath))
            {
                Document tempDoc = new Document();
                DocumentBuilder builder = new DocumentBuilder(tempDoc);
                builder.Writeln("This is a sample PDF generated for conversion.");
                tempDoc.Save(pdfPath, SaveFormat.Pdf);
                Console.WriteLine($"Created placeholder PDF at: {pdfPath}");
            }

            // Ensure the output directory exists.
            string outputDir = Path.GetDirectoryName(xpsPath);
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Load the PDF document.
            Document pdfDocument = new Document(pdfPath);

            // Set XPS save options (optional customizations can be added here).
            XpsSaveOptions xpsOptions = new XpsSaveOptions();

            // Save the document as XPS.
            pdfDocument.Save(xpsPath, xpsOptions);

            Console.WriteLine($"PDF has been successfully converted to XPS at: {xpsPath}");
        }
    }
}
