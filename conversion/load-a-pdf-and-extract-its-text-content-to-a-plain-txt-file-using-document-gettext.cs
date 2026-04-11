using System;
using System.IO;
using System.Text;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for the sample files.
        string dataDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(dataDir);

        // Paths for the input PDF and the output TXT.
        string pdfPath = Path.Combine(dataDir, "sample.pdf");
        string txtPath = Path.Combine(dataDir, "sample.txt");

        // -----------------------------------------------------------------
        // Create a sample PDF document.
        // -----------------------------------------------------------------
        Document createDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(createDoc);
        builder.Writeln("Hello Aspose.Words PDF!");
        createDoc.Save(pdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // Load the PDF and extract its text using Document.GetText().
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);
        string extractedText = pdfDoc.GetText();

        // Save the extracted text to a plain TXT file.
        File.WriteAllText(txtPath, extractedText, Encoding.UTF8);

        // Verify that the TXT file was created.
        if (!File.Exists(txtPath))
            throw new InvalidOperationException("The text file was not created.");

        // Optional: indicate completion (no interactive input required).
        Console.WriteLine("Text extraction completed successfully.");
    }
}
