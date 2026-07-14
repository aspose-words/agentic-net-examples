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
        builder.Writeln("This is a sample PDF file.");
        builder.Writeln("It contains multiple lines of text.");
        source.Save("input.pdf", SaveFormat.Pdf);

        // Load the PDF document.
        Document pdfDoc = new Document("input.pdf");

        // Extract the text content.
        string extractedText = pdfDoc.GetText();

        // Save the extracted text to a plain TXT file.
        File.WriteAllText("output.txt", extractedText);

        // Verify that the output file was created.
        if (!File.Exists("output.txt"))
            throw new InvalidOperationException("The expected output TXT file was not created.");
    }
}
