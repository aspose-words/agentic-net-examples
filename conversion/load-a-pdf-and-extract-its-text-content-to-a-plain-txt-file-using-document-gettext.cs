using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample PDF file.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Hello Aspose.Words PDF extraction example.");
        source.Save("input.pdf", SaveFormat.Pdf);

        if (!File.Exists("input.pdf"))
            throw new InvalidOperationException("Input PDF was not created.");

        // Load the PDF document.
        Document pdfDoc = new Document("input.pdf");

        // Extract text from the PDF.
        string extractedText = pdfDoc.GetText();

        // Save the extracted text to a TXT file.
        File.WriteAllText("output.txt", extractedText);

        if (!File.Exists("output.txt"))
            throw new InvalidOperationException("Output TXT was not created.");

        if (new FileInfo("output.txt").Length == 0)
            throw new InvalidOperationException("Extracted text file is empty.");
    }
}
