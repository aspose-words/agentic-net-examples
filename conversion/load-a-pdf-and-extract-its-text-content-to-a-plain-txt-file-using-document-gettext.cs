using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define file names.
        const string inputPdfPath = "input.pdf";
        const string outputTxtPath = "output.txt";

        // -----------------------------------------------------------------
        // Create a sample PDF document if it does not already exist.
        // -----------------------------------------------------------------
        if (!File.Exists(inputPdfPath))
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln("This is a sample PDF created for text extraction.");
            // Save the document as PDF.
            sampleDoc.Save(inputPdfPath, SaveFormat.Pdf);
        }

        // -----------------------------------------------------------------
        // Load the PDF document.
        // -----------------------------------------------------------------
        Document pdfDocument = new Document(inputPdfPath);

        // -----------------------------------------------------------------
        // Extract the plain text from the PDF.
        // -----------------------------------------------------------------
        string extractedText = pdfDocument.GetText();

        // -----------------------------------------------------------------
        // Write the extracted text to a TXT file.
        // -----------------------------------------------------------------
        File.WriteAllText(outputTxtPath, extractedText);

        // -----------------------------------------------------------------
        // Validate that the output file was created and contains data.
        // -----------------------------------------------------------------
        if (!File.Exists(outputTxtPath) || new FileInfo(outputTxtPath).Length == 0)
        {
            throw new InvalidOperationException("The text extraction failed; output file was not created or is empty.");
        }

        // Optional: indicate success (no console interaction required).
        // Console.WriteLine("Text extraction completed successfully.");
    }
}
