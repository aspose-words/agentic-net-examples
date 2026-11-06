using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define folders for input RTF files and output PDFs.
        string inputFolder = "InputRtf";
        string outputFolder = "OutputPdf";

        // Ensure the input and output directories exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Seed sample RTF files if the input folder is empty.
        string[] existingRtfFiles = Directory.GetFiles(inputFolder, "*.rtf");
        if (existingRtfFiles.Length == 0)
        {
            for (int i = 1; i <= 2; i++)
            {
                // Create a simple document.
                Document sampleDoc = new Document();
                DocumentBuilder builder = new DocumentBuilder(sampleDoc);
                builder.Writeln($"Sample RTF content for file {i}.");

                // Save as RTF.
                string rtfPath = Path.Combine(inputFolder, $"Sample{i}.rtf");
                sampleDoc.Save(rtfPath, SaveFormat.Rtf);
            }
        }

        // Process each RTF file in the input folder.
        foreach (string rtfFilePath in Directory.GetFiles(inputFolder, "*.rtf"))
        {
            // Load the RTF document.
            Document doc = new Document(rtfFilePath);

            // Determine the output PDF path.
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(rtfFilePath);
            string pdfPath = Path.Combine(outputFolder, $"{fileNameWithoutExt}.pdf");

            // Save the document as PDF using default layout.
            doc.Save(pdfPath, SaveFormat.Pdf);

            // Validate that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Expected PDF was not created: {pdfPath}");
        }
    }
}
