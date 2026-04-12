using System;
using System.IO;
using Aspose.Words;

public class BatchRtfToPdfConverter
{
    public static void Main()
    {
        // Define input and output directories relative to the current working directory.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputRtf");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputPdf");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a few sample RTF files if the folder is empty.
        string[] sampleFiles = Directory.GetFiles(inputFolder, "*.rtf");
        if (sampleFiles.Length == 0)
        {
            for (int i = 1; i <= 2; i++)
            {
                Document sampleDoc = new Document();
                DocumentBuilder builder = new DocumentBuilder(sampleDoc);
                builder.Writeln($"This is sample RTF document #{i}.");
                string rtfPath = Path.Combine(inputFolder, $"Sample{i}.rtf");
                sampleDoc.Save(rtfPath, SaveFormat.Rtf);
            }
        }

        // Process each RTF file in the input folder.
        foreach (string rtfFilePath in Directory.GetFiles(inputFolder, "*.rtf"))
        {
            // Load the RTF document.
            Document doc = new Document(rtfFilePath);

            // Determine the output PDF file path.
            string pdfFileName = Path.GetFileNameWithoutExtension(rtfFilePath) + ".pdf";
            string pdfFilePath = Path.Combine(outputFolder, pdfFileName);

            // Save the document as PDF using the default layout.
            doc.Save(pdfFilePath, SaveFormat.Pdf);

            // Verify that the PDF file was created.
            if (!File.Exists(pdfFilePath))
            {
                throw new InvalidOperationException($"Failed to create PDF file: {pdfFilePath}");
            }

            Console.WriteLine($"Converted '{Path.GetFileName(rtfFilePath)}' to '{pdfFileName}'.");
        }

        Console.WriteLine("Batch conversion completed.");
    }
}
