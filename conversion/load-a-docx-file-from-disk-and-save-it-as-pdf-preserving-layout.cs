using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file paths relative to the executable directory.
        string inputPath = Path.Combine(Environment.CurrentDirectory, "Sample.docx");
        string outputPath = Path.Combine(Environment.CurrentDirectory, "Sample.pdf");

        // -----------------------------------------------------------------
        // Step 1: Create a sample DOCX document.
        // -----------------------------------------------------------------
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("Hello Aspose.Words!");
        // Save the document as DOCX so it can be loaded later.
        sampleDoc.Save(inputPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Step 2: Load the DOCX document from disk.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(inputPath);

        // -----------------------------------------------------------------
        // Step 3: Convert and save the document as PDF, preserving layout.
        // -----------------------------------------------------------------
        loadedDoc.Save(outputPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // Step 4: Verify that the PDF file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The PDF conversion failed; output file not found.", outputPath);
        }

        Console.WriteLine($"Conversion successful. PDF saved to: {outputPath}");
    }
}
