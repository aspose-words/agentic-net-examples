using System;
using System.IO;
using Aspose.Words;

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
        researchBuilder.Writeln("Author: John Doe");
        researchBuilder.Writeln();
        researchBuilder.Writeln("Introduction");
        researchBuilder.Writeln("This is the introduction section of the research paper.");
        // Insert a simple PAGE field to demonstrate field updating later.
        researchBuilder.InsertField("PAGE  \\* MERGEFORMAT");
        researchDoc.Save(researchPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Create a sample bibliography document.
        // -----------------------------------------------------------------
        Document bibliographyDoc = new Document();
        DocumentBuilder bibBuilder = new DocumentBuilder(bibliographyDoc);
        bibBuilder.Writeln("Bibliography");
        // Insert a BIBLIOGRAPHY field; it will be updated after merging.
        bibBuilder.InsertField("BIBLIOGRAPHY");
        bibliographyDoc.Save(bibliographyPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Load the created documents.
        // -----------------------------------------------------------------
        Document research = new Document(researchPath);
        Document bibliography = new Document(bibliographyPath);

        // -----------------------------------------------------------------
        // Append the bibliography to the research paper.
        // Keep source formatting to preserve the bibliography style.
        // -----------------------------------------------------------------
        research.AppendDocument(bibliography, ImportFormatMode.KeepSourceFormatting);

        // -----------------------------------------------------------------
        // Update all fields (including PAGE and BIBLIOGRAPHY).
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
    }
}
