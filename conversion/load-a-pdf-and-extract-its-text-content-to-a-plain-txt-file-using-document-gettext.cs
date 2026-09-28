using System;
using System.IO;
using Aspose.Words;

public class PdfToTextExtractor
{
    public static void Main()
    {
        // Create a sample PDF file.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("This is a sample PDF document.");
        builder.Writeln("It contains multiple lines of text.");
        string pdfPath = "sample.pdf";
        sampleDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF document.
        Document pdfDocument = new Document(pdfPath);

        // Extract text from the PDF.
        string extractedText = pdfDocument.GetText();

        // Save the extracted text to a TXT file.
        string txtPath = "output.txt";
        File.WriteAllText(txtPath, extractedText);

        // Validate that the TXT file was created and contains data.
        if (!File.Exists(txtPath) || new FileInfo(txtPath).Length == 0)
        {
            throw new InvalidOperationException("The text extraction failed; output file was not created or is empty.");
        }

        // Optionally, clean up the sample PDF (comment out if you want to keep it).
        // File.Delete(pdfPath);
    }
}
