using System;
using System.IO;
using Aspose.Words;

namespace PdfToDocxConversion
{
    class Program
    {
        static void Main(string[] args)
        {
            // Use paths relative to the executable directory so the example works out‑of‑the‑box.
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string pdfPath = Path.Combine(baseDir, "Input", "form_fields.pdf");
            string docxPath = Path.Combine(baseDir, "Output", "form_fields_converted.docx");

            // Ensure the output directory exists.
            Directory.CreateDirectory(Path.GetDirectoryName(docxPath)!);

            if (!File.Exists(pdfPath))
            {
                Console.WriteLine($"Input PDF not found at: {pdfPath}");
                Console.WriteLine("Place a PDF file named 'form_fields.pdf' in the 'Input' folder and rerun the program.");
                return;
            }

            // Load the PDF document. Aspose.Words automatically detects the format.
            // Form fields are imported as Word form fields (StructuredDocumentTag objects).
            Document pdfDocument = new Document(pdfPath);

            // Save the loaded document as DOCX. The file extension determines the save format.
            pdfDocument.Save(docxPath);

            Console.WriteLine($"Conversion succeeded. DOCX saved to: {docxPath}");
        }
    }
}
