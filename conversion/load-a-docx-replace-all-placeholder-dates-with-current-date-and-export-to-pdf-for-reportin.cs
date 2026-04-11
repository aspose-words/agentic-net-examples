using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define directories and file names.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        string inputPath = Path.Combine(artifactsDir, "SampleDocument.docx");
        string outputPath = Path.Combine(artifactsDir, "Report.pdf");

        // -----------------------------------------------------------------
        // 1. Create a sample DOCX file with a placeholder for the date.
        // -----------------------------------------------------------------
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("Monthly Report");
        builder.Writeln("Generated on {{Date}}."); // Placeholder to be replaced.
        sampleDoc.Save(inputPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Load the DOCX file.
        // -----------------------------------------------------------------
        Document doc = new Document(inputPath);

        // -----------------------------------------------------------------
        // 3. Replace all date placeholders with the current date.
        // -----------------------------------------------------------------
        string placeholder = "{{Date}}";
        string currentDate = DateTime.Now.ToString("yyyy-MM-dd");
        doc.Range.Replace(placeholder, currentDate);

        // -----------------------------------------------------------------
        // 4. Convert the document to PDF.
        // -----------------------------------------------------------------
        doc.Save(outputPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 5. Verify that the PDF was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The PDF file was not created.", outputPath);
        }

        // Optional: indicate success (no interactive input required).
        Console.WriteLine("PDF report generated successfully at: " + outputPath);
    }
}
