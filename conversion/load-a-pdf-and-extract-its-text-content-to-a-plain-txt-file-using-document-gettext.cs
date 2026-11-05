using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main(string[] args)
    {
        // Create a sample PDF file as input.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample PDF document.");
        sourceDoc.Save("input.pdf", SaveFormat.Pdf);

        // Load the PDF document.
        Document pdfDoc = new Document("input.pdf");

        // Extract text from the PDF.
        string extractedText = pdfDoc.GetText();

        // Save the extracted text to a TXT file.
        string outputPath = "output.txt";
        File.WriteAllText(outputPath, extractedText);

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("Expected output TXT file was not created.");
        }

        // Optionally, you could verify that the file is not empty.
        if (new FileInfo(outputPath).Length == 0)
        {
            throw new InvalidOperationException("The extracted text file is empty.");
        }
    }
}
