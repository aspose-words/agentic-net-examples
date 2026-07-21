using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class BatchDocToPdfConverter
{
    public static void Main()
    {
        // Prepare input folder and create sample DOC files.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        Directory.CreateDirectory(inputFolder);

        for (int i = 1; i <= 3; i++)
        {
            string docPath = Path.Combine(inputFolder, $"Sample{i}.doc");

            // Create a blank document and add sample content.
            Document source = new Document();
            DocumentBuilder builder = new DocumentBuilder(source);
            builder.Writeln($"Sample DOC content {i}.");

            // Save the document as DOC.
            source.Save(docPath, SaveFormat.Doc);
        }

        // Prepare output folder for PDFs.
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputPdfs");
        Directory.CreateDirectory(outputFolder);

        // Process each DOC file in the input folder.
        foreach (string docFile in Directory.GetFiles(inputFolder, "*.doc"))
        {
            // Load the DOC file.
            Document doc = new Document(docFile);

            // Determine the PDF output path.
            string pdfFileName = Path.GetFileNameWithoutExtension(docFile) + ".pdf";
            string pdfPath = Path.Combine(outputFolder, pdfFileName);

            // Convert and save as PDF.
            doc.Save(pdfPath, SaveFormat.Pdf);

            // Verify that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Expected PDF was not created for '{docFile}'.");

            // Log conversion status.
            Console.WriteLine($"Converted '{docFile}' to '{pdfPath}'.");
        }
    }
}
