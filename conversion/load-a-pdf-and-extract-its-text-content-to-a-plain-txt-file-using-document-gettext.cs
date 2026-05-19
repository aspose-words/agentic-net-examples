using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample PDF file.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Hello Aspose.Words PDF extraction.");
        const string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Expected PDF file was not created.");

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Extract plain text from the PDF.
        string extractedText = pdfDoc.GetText();

        // Save the extracted text to a TXT file.
        const string txtPath = "output.txt";
        File.WriteAllText(txtPath, extractedText);
        if (!File.Exists(txtPath))
            throw new InvalidOperationException("Expected TXT file was not created.");
    }
}
