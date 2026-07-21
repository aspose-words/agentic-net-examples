using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample PDF document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample PDF created for text extraction.");
        string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Extract the text content.
        string extractedText = pdfDoc.GetText();

        // Save the extracted text to a plain TXT file.
        string txtPath = "extracted.txt";
        File.WriteAllText(txtPath, extractedText);

        // Validate that the TXT file was created.
        if (!File.Exists(txtPath))
            throw new InvalidOperationException("The output TXT file was not created.");

        // Optional: indicate success (no interactive output required).
        Console.WriteLine("Text extraction completed successfully.");
    }
}
