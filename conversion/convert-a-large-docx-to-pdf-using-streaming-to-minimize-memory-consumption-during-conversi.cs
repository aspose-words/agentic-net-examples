using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample large DOCX document.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        // Simulate a large document by writing many lines.
        for (int i = 0; i < 5000; i++)
        {
            builder.Writeln($"This is line {i + 1} of a large document.");
        }
        source.Save("input.docx", SaveFormat.Docx);

        // Load the DOCX document.
        Document doc = new Document("input.docx");

        // Create PDF save options with memory optimization enabled.
        SaveOptions saveOptions = SaveOptions.CreateSaveOptions(SaveFormat.Pdf);
        saveOptions.MemoryOptimization = true;

        // Save the document to a memory stream to minimize memory usage.
        using (MemoryStream pdfStream = new MemoryStream())
        {
            doc.Save(pdfStream, saveOptions);

            // Verify that data was written to the stream.
            if (pdfStream.Length == 0)
                throw new InvalidOperationException("No PDF data was written to the memory stream.");

            // Write the stream contents to a PDF file.
            pdfStream.Position = 0;
            using (FileStream file = new FileStream("output.pdf", FileMode.Create, FileAccess.Write))
            {
                pdfStream.CopyTo(file);
            }
        }

        // Validate that the output PDF file was created.
        if (!File.Exists("output.pdf"))
            throw new InvalidOperationException("The expected output PDF file was not created.");
    }
}
