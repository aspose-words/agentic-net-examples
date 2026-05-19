using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define folders for input DOC files and output PDF files.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputPdfs");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Seed the input folder with sample DOC files if it is empty.
        string[] existingDocs = Directory.GetFiles(inputFolder, "*.doc");
        if (existingDocs.Length == 0)
        {
            for (int i = 1; i <= 3; i++)
            {
                // Create a blank document and add sample content.
                Document source = new Document();
                DocumentBuilder builder = new DocumentBuilder(source);
                builder.Writeln($"Sample DOC content {i}.");

                // Save the document as DOC.
                string docPath = Path.Combine(inputFolder, $"sample{i}.doc");
                source.Save(docPath, SaveFormat.Doc);
            }
        }

        // Process each DOC file in the input folder.
        string[] docFiles = Directory.GetFiles(inputFolder, "*.doc");
        foreach (string docFile in docFiles)
        {
            try
            {
                // Load the DOC file.
                Document doc = new Document(docFile);

                // Determine the output PDF path.
                string pdfFileName = Path.GetFileNameWithoutExtension(docFile) + ".pdf";
                string pdfPath = Path.Combine(outputFolder, pdfFileName);

                // Convert and save as PDF.
                doc.Save(pdfPath, SaveFormat.Pdf);

                // Verify that the PDF was created.
                if (!File.Exists(pdfPath))
                    throw new InvalidOperationException($"PDF was not created for '{docFile}'.");

                // Log successful conversion.
                Console.WriteLine($"Converted '{Path.GetFileName(docFile)}' to PDF: '{pdfFileName}'.");
            }
            catch (Exception ex)
            {
                // Log any errors that occur during conversion.
                Console.WriteLine($"Error converting '{Path.GetFileName(docFile)}': {ex.Message}");
            }
        }

        // Optional: indicate completion.
        Console.WriteLine("Batch conversion completed.");
    }
}
