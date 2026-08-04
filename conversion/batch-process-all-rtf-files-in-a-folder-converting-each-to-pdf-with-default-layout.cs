using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define folders for input RTF files and output PDFs.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputRtf");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputPdf");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample RTF files if the input folder is empty.
        if (Directory.GetFiles(inputFolder, "*.rtf").Length == 0)
        {
            for (int i = 1; i <= 3; i++)
            {
                // Create a blank document and add some text.
                Document sampleDoc = new Document();
                DocumentBuilder builder = new DocumentBuilder(sampleDoc);
                builder.Writeln($"Sample RTF content for file {i}.");

                // Save the document as RTF in the input folder.
                string rtfPath = Path.Combine(inputFolder, $"Sample{i}.rtf");
                sampleDoc.Save(rtfPath, SaveFormat.Rtf);
            }
        }

        // Process each RTF file in the input folder.
        string[] rtfFiles = Directory.GetFiles(inputFolder, "*.rtf");
        foreach (string rtfFile in rtfFiles)
        {
            // Load the RTF document.
            Document doc = new Document(rtfFile);

            // Determine the output PDF path.
            string pdfFileName = Path.GetFileNameWithoutExtension(rtfFile) + ".pdf";
            string pdfPath = Path.Combine(outputFolder, pdfFileName);

            // Convert and save as PDF using the default layout.
            doc.Save(pdfPath, SaveFormat.Pdf);

            // Verify that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create PDF: {pdfPath}");
        }
    }
}
