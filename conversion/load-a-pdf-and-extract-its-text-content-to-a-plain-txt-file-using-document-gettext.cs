using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample PDF document.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Sample PDF content for text extraction.");
        source.Save("input.pdf", SaveFormat.Pdf);

        // Load the PDF document.
        Document pdfDoc = new Document("input.pdf");

        // Extract the text content.
        string extractedText = pdfDoc.GetText();

        // Save the extracted text to a plain TXT file.
        File.WriteAllText("output.txt", extractedText);

        // Validate that the TXT file was created and contains data.
        if (!File.Exists("output.txt"))
            throw new InvalidOperationException("Expected output TXT was not created.");

        if (new FileInfo("output.txt").Length == 0)
            throw new InvalidOperationException("The extracted text file is empty.");
    }
}
