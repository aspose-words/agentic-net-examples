using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Define output directory and file paths.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        string researchPath = Path.Combine(outputDir, "ResearchPaper.docx");
        string bibliographyPath = Path.Combine(outputDir, "Bibliography.docx");
        string mergedPath = Path.Combine(outputDir, "MergedDocument.docx");
        string pdfPath = Path.Combine(outputDir, "MergedDocument.pdf");

        // -----------------------------------------------------------------
        // Create a sample research paper document.
        // -----------------------------------------------------------------
        Document researchDoc = new Document();
        DocumentBuilder researchBuilder = new DocumentBuilder(researchDoc);
        researchBuilder.Writeln("Research Paper Title");
        researchBuilder.Writeln("Introduction");
        researchBuilder.Writeln("This is the content of the research paper.");
        researchDoc.Save(researchPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Create a sample bibliography document containing a BIBLIOGRAPHY field.
        // -----------------------------------------------------------------
        Document bibDoc = new Document();
        DocumentBuilder bibBuilder = new DocumentBuilder(bibDoc);
        bibBuilder.Writeln("Bibliography");
        // Insert a bibliography field (will be updated later).
        bibBuilder.InsertField(FieldType.FieldBibliography, true);
        bibDoc.Save(bibliographyPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Load the two documents.
        // -----------------------------------------------------------------
        Document research = new Document(researchPath);
        Document bibliography = new Document(bibliographyPath);

        // -----------------------------------------------------------------
        // Append the bibliography document to the research paper.
        // Keep the source formatting to preserve the bibliography field.
        // -----------------------------------------------------------------
        research.AppendDocument(bibliography, ImportFormatMode.KeepSourceFormatting);

        // -----------------------------------------------------------------
        // Update all fields (including the bibliography field) in the merged document.
        // -----------------------------------------------------------------
        research.UpdateFields();

        // -----------------------------------------------------------------
        // Save the merged document as DOCX.
        // -----------------------------------------------------------------
        research.Save(mergedPath, SaveFormat.Docx);

        // Validate that the merged DOCX was created.
        if (!File.Exists(mergedPath))
            throw new InvalidOperationException("Merged DOCX file was not created.");

        // -----------------------------------------------------------------
        // Export the merged document to PDF.
        // -----------------------------------------------------------------
        research.Save(pdfPath, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // Optional: write status messages to the console.
        Console.WriteLine("Documents merged and saved successfully:");
        Console.WriteLine($"- Merged DOCX: {mergedPath}");
        Console.WriteLine($"- PDF: {pdfPath}");
    }
}
