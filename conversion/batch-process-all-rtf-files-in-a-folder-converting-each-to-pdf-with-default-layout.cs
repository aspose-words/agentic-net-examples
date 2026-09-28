using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define the folder that will contain the sample RTF files.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputRtf");
        Directory.CreateDirectory(inputFolder);

        // Create sample RTF files.
        CreateSampleRtf(Path.Combine(inputFolder, "Sample1.rtf"), "This is the first sample RTF document.");
        CreateSampleRtf(Path.Combine(inputFolder, "Sample2.rtf"), "This is the second sample RTF document.");

        // Process each RTF file in the folder and convert it to PDF.
        foreach (string rtfPath in Directory.GetFiles(inputFolder, "*.rtf"))
        {
            // Load the RTF document.
            Document doc = new Document(rtfPath);

            // Determine the output PDF path.
            string pdfPath = Path.ChangeExtension(rtfPath, ".pdf");

            // Save the document as PDF using default layout.
            doc.Save(pdfPath, SaveFormat.Pdf);

            // Validate that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"PDF file was not created: {pdfPath}");
        }
    }

    private static void CreateSampleRtf(string filePath, string content)
    {
        // Create a new empty document.
        Document doc = new Document();

        // Add content to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(content);

        // Save the document as RTF.
        doc.Save(filePath, SaveFormat.Rtf);

        // Validate that the RTF file was created.
        if (!File.Exists(filePath))
            throw new InvalidOperationException($"RTF file was not created: {filePath}");
    }
}
