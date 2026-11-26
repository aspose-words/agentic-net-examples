using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare input and output folders.
        string inputFolder = "InputRtf";
        string outputFolder = "OutputPdf";

        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample RTF files if the folder is empty.
        string[] sampleFiles = { "Sample1.rtf", "Sample2.rtf" };
        foreach (string fileName in sampleFiles)
        {
            string inputPath = Path.Combine(inputFolder, fileName);
            if (!File.Exists(inputPath))
            {
                // Create a simple document and save it as RTF.
                Document source = new Document();
                DocumentBuilder builder = new DocumentBuilder(source);
                builder.Writeln($"This is the content of {Path.GetFileNameWithoutExtension(fileName)}.");
                source.Save(inputPath, SaveFormat.Rtf);
            }
        }

        // Process each RTF file in the input folder.
        string[] rtfFiles = Directory.GetFiles(inputFolder, "*.rtf");
        foreach (string rtfPath in rtfFiles)
        {
            // Load the RTF document.
            Document doc = new Document(rtfPath);

            // Determine the output PDF path.
            string pdfFileName = Path.GetFileNameWithoutExtension(rtfPath) + ".pdf";
            string pdfPath = Path.Combine(outputFolder, pdfFileName);

            // Save as PDF using default layout.
            doc.Save(pdfPath, SaveFormat.Pdf);

            // Validate that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"PDF file was not created: {pdfPath}");
        }

        // Optional: indicate completion (no interactive wait).
        Console.WriteLine("Batch conversion completed successfully.");
    }
}
