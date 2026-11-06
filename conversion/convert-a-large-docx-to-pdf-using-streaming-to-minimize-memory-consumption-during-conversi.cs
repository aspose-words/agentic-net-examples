using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample large DOCX document.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        // Simulate a large document by writing many lines.
        for (int i = 0; i < 5000; i++)
        {
            builder.Writeln($"This is line {i + 1} of a large document.");
        }

        // Save the source document as DOCX.
        string inputPath = "input.docx";
        source.Save(inputPath, SaveFormat.Docx);

        // Load the DOCX document.
        Document doc = new Document(inputPath);

        // Convert to PDF using a FileStream to stream the output.
        string outputPath = "output.pdf";
        using (FileStream pdfStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
        {
            doc.Save(pdfStream, SaveFormat.Pdf);
        }

        // Validate that the PDF was created and contains data.
        if (!File.Exists(outputPath) || new FileInfo(outputPath).Length == 0)
        {
            throw new InvalidOperationException("Expected output PDF was not created or is empty.");
        }
    }
}
