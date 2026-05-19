using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Define file paths for the sample input and the output document.
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "Sample.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Watermarked.docx");

        // -----------------------------------------------------------------
        // Step 1: Create a simple source document and save it to the input path.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample document created for the watermark example.");
        sourceDoc.Save(inputPath);

        // -----------------------------------------------------------------
        // Step 2: Load the document from the file system.
        // -----------------------------------------------------------------
        Document doc = new Document(inputPath);

        // -----------------------------------------------------------------
        // Step 3: Add a text watermark using the Document.Watermark API.
        // -----------------------------------------------------------------
        doc.Watermark.SetText("CONFIDENTIAL");

        // -----------------------------------------------------------------
        // Step 4: Save the watermarked document.
        // -----------------------------------------------------------------
        doc.Save(outputPath);
    }
}
