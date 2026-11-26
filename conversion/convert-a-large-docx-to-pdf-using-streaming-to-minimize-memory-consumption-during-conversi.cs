using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a large DOCX file as sample input.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        for (int i = 0; i < 5000; i++)
        {
            builder.Writeln($"This is line {i + 1} of a large document.");
        }
        string inputPath = "large_input.docx";
        source.Save(inputPath, SaveFormat.Docx);

        // Verify the input file was created.
        if (!File.Exists(inputPath))
            throw new InvalidOperationException("Input DOCX file was not created.");

        // Step 2: Load the DOCX document.
        Document doc = new Document(inputPath);

        // Step 3: Convert to PDF using a streaming approach (FileStream).
        string outputPath = "output.pdf";
        using (FileStream pdfStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
        {
            doc.Save(pdfStream, SaveFormat.Pdf);
        }

        // Step 4: Validate that the PDF was created and contains data.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Expected output PDF was not created.");

        FileInfo pdfInfo = new FileInfo(outputPath);
        if (pdfInfo.Length == 0)
            throw new InvalidOperationException("The generated PDF file is empty.");

        // Optional: Clean up sample files (comment out if you want to keep them).
        // File.Delete(inputPath);
        // File.Delete(outputPath);
    }
}
