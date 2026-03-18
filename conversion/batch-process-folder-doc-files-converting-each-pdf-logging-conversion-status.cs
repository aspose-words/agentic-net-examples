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
        // Adjust these paths as needed.
        private const string InputFolder = @"C:\Docs\Input";
        private const string OutputFolder = @"C:\Docs\Output";

        public static void Main()
        {
            // Ensure the output directory exists.
            Directory.CreateDirectory(OutputFolder);

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
