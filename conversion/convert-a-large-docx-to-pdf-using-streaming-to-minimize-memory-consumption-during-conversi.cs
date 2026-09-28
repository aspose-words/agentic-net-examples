using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a large DOCX document locally.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        for (int i = 0; i < 5000; i++)
        {
            builder.Writeln($"Paragraph {i + 1}: This is sample text to increase document size.");
        }
        string inputPath = "large_input.docx";
        source.Save(inputPath, SaveFormat.Docx);

        // Load the DOCX document.
        Document doc = new Document(inputPath);

        // Convert to PDF using a file stream to minimize memory usage.
        string outputPath = "output.pdf";
        using (FileStream pdfStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
        {
            doc.Save(pdfStream, SaveFormat.Pdf);
        }

        // Validate that the PDF was created and contains data.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("Expected output PDF was not created.");
        }

        FileInfo pdfInfo = new FileInfo(outputPath);
        if (pdfInfo.Length == 0)
        {
            throw new InvalidOperationException("The generated PDF file is empty.");
        }
    }
}
