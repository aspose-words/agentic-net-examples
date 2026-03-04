using System;
using System.IO;
using Aspose.Words;

namespace ExtractionDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the DOCX file to be extracted.
            string docxPath = @"C:\Docs\SampleDocument.docx";

            // Ensure the file exists before attempting extraction.
            if (!File.Exists(docxPath))
            {
                Console.WriteLine($"File not found: {docxPath}");
                return;
            }

            try
            {
                // Load the document as plain text using the PlainTextDocument constructor.
                // This automatically detects the file format.
                PlainTextDocument plainTextDoc = new PlainTextDocument(docxPath);

                // Retrieve the concatenated textual content of the document.
                string extractedText = plainTextDoc.Text;

                // Output the extracted text to the console (or process as needed).
                Console.WriteLine("=== Extracted Text Start ===");
                Console.WriteLine(extractedText);
                Console.WriteLine("=== Extracted Text End ===");
            }
            catch (UnsupportedFileFormatException ex)
            {
                Console.WriteLine($"Unsupported format: {ex.Message}");
            }
            catch (FileCorruptedException ex)
            {
                Console.WriteLine($"File is corrupted: {ex.Message}");
            }
            catch (IncorrectPasswordException ex)
            {
                Console.WriteLine($"Document is encrypted and requires a password: {ex.Message}");
            }
            catch (Exception ex)
            {
                // Catch any other unexpected exceptions.
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
