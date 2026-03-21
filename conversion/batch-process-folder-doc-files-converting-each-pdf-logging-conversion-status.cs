using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsBatchConversion
{
    /// <summary>
    /// Demonstrates batch conversion of all *.doc files in a folder to PDF.
    /// Uses Aspose.Words Document loading and saving APIs as defined in the provided rules.
    /// </summary>
    public static class Program
    {
        // Use folders relative to the executable so they always exist.
        private static readonly string BaseFolder = AppContext.BaseDirectory;
        private static readonly string InputFolder = Path.Combine(BaseFolder, "Input");
        private static readonly string OutputFolder = Path.Combine(BaseFolder, "Output");

        public static void Main()
        {
            // Ensure the input and output directories exist.
            Directory.CreateDirectory(InputFolder);
            Directory.CreateDirectory(OutputFolder);

            // If there are no .doc files, create a simple sample document to demonstrate conversion.
            if (Directory.GetFiles(InputFolder, "*.doc", SearchOption.TopDirectoryOnly).Length == 0)
            {
                var sampleDoc = new Document();
                var builder = new DocumentBuilder(sampleDoc);
                builder.Writeln("This is a sample document generated at " + DateTime.Now);
                string samplePath = Path.Combine(InputFolder, "Sample.doc");
                sampleDoc.Save(samplePath, SaveFormat.Doc);
                Console.WriteLine($"[INFO] Created sample document at '{samplePath}'.");
            }

            // Get all files with .doc extension (including .docx if desired).
            string[] docFiles = Directory.GetFiles(InputFolder, "*.doc", SearchOption.TopDirectoryOnly);

            foreach (string docPath in docFiles)
            {
                try
                {
                    // Load the source Word document using the Document(string) constructor.
                    Document doc = new Document(docPath);

                    // Build the output PDF file name.
                    string pdfFileName = Path.GetFileNameWithoutExtension(docPath) + ".pdf";
                    string pdfPath = Path.Combine(OutputFolder, pdfFileName);

                    // Save the document as PDF using the Save(string, SaveFormat) overload.
                    doc.Save(pdfPath, SaveFormat.Pdf);

                    // Log successful conversion.
                    Console.WriteLine($"[SUCCESS] Converted '{docPath}' to '{pdfPath}'.");
                }
                catch (Exception ex)
                {
                    // Log any errors that occur during conversion.
                    Console.WriteLine($"[ERROR] Failed to convert '{docPath}'. Exception: {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
