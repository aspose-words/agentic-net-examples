using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class BatchDocToPdfConverter
{
    public static void Main()
    {
        // Prepare input and output directories.
        string currentDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(currentDir, "InputDocs");
        string outputDir = Path.Combine(currentDir, "OutputPdfs");

        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample DOC files if the folder is empty.
        string[] existingDocs = Directory.GetFiles(inputDir, "*.doc");
        if (existingDocs.Length == 0)
        {
            for (int i = 1; i <= 3; i++)
            {
                Document sampleDoc = new Document();
                DocumentBuilder builder = new DocumentBuilder(sampleDoc);
                builder.Writeln($"This is sample document #{i}.");
                string docPath = Path.Combine(inputDir, $"Sample{i}.doc");
                sampleDoc.Save(docPath, SaveFormat.Doc);
            }
        }

        // Process each DOC file in the input folder.
        string[] docFiles = Directory.GetFiles(inputDir, "*.doc");
        foreach (string docFilePath in docFiles)
        {
            try
            {
                // Load the DOC file.
                Document doc = new Document(docFilePath);

                // Determine the PDF output path.
                string pdfFileName = Path.GetFileNameWithoutExtension(docFilePath) + ".pdf";
                string pdfFilePath = Path.Combine(outputDir, pdfFileName);

                // Convert and save as PDF.
                doc.Save(pdfFilePath, SaveFormat.Pdf);

                // Verify that the PDF was created.
                if (!File.Exists(pdfFilePath))
                {
                    throw new InvalidOperationException($"PDF file was not created: {pdfFilePath}");
                }

                Console.WriteLine($"Converted '{Path.GetFileName(docFilePath)}' to PDF successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error converting '{Path.GetFileName(docFilePath)}': {ex.Message}");
            }
        }

        // List the resulting PDF files.
        Console.WriteLine("Conversion finished. Generated PDF files:");
        foreach (string pdfPath in Directory.GetFiles(outputDir, "*.pdf"))
        {
            Console.WriteLine($" - {Path.GetFileName(pdfPath)}");
        }
    }
}
