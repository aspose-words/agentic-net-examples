using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define folders for input RTF files and output PDF files.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputRtf");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputPdf");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample RTF documents if the input folder is empty.
        // This follows the "doc to pdf file pattern" adapted for RTF.
        string[] sampleNames = { "Sample1.rtf", "Sample2.rtf", "Sample3.rtf" };
        foreach (string fileName in sampleNames)
        {
            string filePath = Path.Combine(inputFolder, fileName);
            if (!File.Exists(filePath))
            {
                Document source = new Document();
                DocumentBuilder builder = new DocumentBuilder(source);
                builder.Writeln($"This is the content of {Path.GetFileNameWithoutExtension(fileName)}.");
                source.Save(filePath, SaveFormat.Rtf);
            }
        }

        // Process each RTF file in the input folder and convert it to PDF.
        string[] rtfFiles = Directory.GetFiles(inputFolder, "*.rtf");
        foreach (string rtfPath in rtfFiles)
        {
            // Load the RTF document.
            Document doc = new Document(rtfPath);

            // Determine the output PDF path.
            string pdfFileName = Path.GetFileNameWithoutExtension(rtfPath) + ".pdf";
            string pdfPath = Path.Combine(outputFolder, pdfFileName);

            // Save the document as PDF using the default layout.
            doc.Save(pdfPath, SaveFormat.Pdf);

            // Verify that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Expected output PDF was not created: {pdfPath}");
        }

        // Optional: indicate completion.
        Console.WriteLine("Batch conversion completed successfully.");
    }
}
