// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class RenderWithStylisticSet
{
    static void Main()
    {
        // Define paths for input and output files.
        string dataDir = @"C:\Data\";
        string inputPath = Path.Combine(dataDir, "input.docx");
        string outputPath = Path.Combine(dataDir, "output.pdf");

        Document doc;

        // Load an existing document if it exists; otherwise create a new one.
        if (File.Exists(inputPath))
        {
            // Load rule
            doc = new Document(inputPath);
        }
        else
        {
            // Create rule
            doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Enable a specific stylistic set (e.g., Stylistic Set 1).
            builder.Font.Name = "Arial";
            builder.Font.StylisticSet = 1; // activates the chosen stylistic set

            // Add sample content.
            builder.Writeln("This text is rendered with Stylistic Set 1 enabled.");
        }

        // Configure PDF save options (create rule).
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save rule: render the document as PDF using the specified options.
        doc.Save(outputPath, pdfOptions);
    }
}
