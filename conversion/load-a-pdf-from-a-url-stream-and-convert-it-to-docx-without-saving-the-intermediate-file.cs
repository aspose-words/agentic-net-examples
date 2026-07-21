using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample Word document in memory.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Sample content for PDF conversion.");

        // Save the document as PDF into a memory stream.
        using (MemoryStream pdfStream = new MemoryStream())
        {
            source.Save(pdfStream, SaveFormat.Pdf);
            pdfStream.Position = 0; // Reset for reading.

            // Load the PDF from the memory stream.
            Document pdfDoc = new Document(pdfStream);

            // Convert the PDF to DOCX and save to disk.
            const string outputPath = "output.docx";
            pdfDoc.Save(outputPath, SaveFormat.Docx);

            // Verify that the DOCX file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The DOCX output file was not created.");
        }

        Console.WriteLine("Conversion completed successfully.");
    }
}
