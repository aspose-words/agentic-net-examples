using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names for the sample documents and the results.
        const string researchPath = "ResearchPaper.docx";
        const string bibliographyPath = "Bibliography.docx";
        const string mergedPath = "MergedResearchPaper.docx";
        const string pdfPath = "MergedResearchPaper.pdf";

        // -----------------------------------------------------------------
        // 1. Create a sample research paper document.
        // -----------------------------------------------------------------
        Document researchDoc = new Document();
        DocumentBuilder researchBuilder = new DocumentBuilder(researchDoc);
        researchBuilder.Writeln("Research Paper Title");
        researchBuilder.Writeln("This is the introduction of the research paper.");
        researchBuilder.Writeln("Main content goes here...");
        // Save the research paper as DOCX.
        researchDoc.Save(researchPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Create a sample bibliography document.
        // -----------------------------------------------------------------
        Document bibliographyDoc = new Document();
        DocumentBuilder bibBuilder = new DocumentBuilder(bibliographyDoc);
        bibBuilder.Writeln("Bibliography");
        // Insert a BIBLIOGRAPHY field to demonstrate field updating.
        bibBuilder.InsertField("BIBLIOGRAPHY");
        // Add a placeholder reference.
        bibBuilder.Writeln("Reference 1: Example Book, Author, 2020.");
        // Save the bibliography as DOCX.
        bibliographyDoc.Save(bibliographyPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 3. Load the created documents.
        // -----------------------------------------------------------------
        Document research = new Document(researchPath);
        Document bibliography = new Document(bibliographyPath);

        // -----------------------------------------------------------------
        // 4. Append the bibliography document to the research paper.
        // -----------------------------------------------------------------
        research.AppendDocument(bibliography, ImportFormatMode.KeepSourceFormatting);

        // -----------------------------------------------------------------
        // 5. Update all fields (e.g., the BIBLIOGRAPHY field).
        // -----------------------------------------------------------------
        research.UpdateFields();

        // -----------------------------------------------------------------
        // 6. Save the merged document in DOCX format.
        // -----------------------------------------------------------------
        research.Save(mergedPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 7. Export the merged document to PDF.
        // -----------------------------------------------------------------
        research.Save(pdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 8. Validation: ensure output files exist and contain expected content.
        // -----------------------------------------------------------------
        if (!File.Exists(mergedPath))
            throw new InvalidOperationException($"Merged DOCX file was not created: {mergedPath}");

        if (!File.Exists(pdfPath))
            throw new InvalidOperationException($"PDF file was not created: {pdfPath}");

        string mergedText = research.GetText();
        if (!mergedText.Contains("Bibliography"))
            throw new InvalidOperationException("Merged document does not contain the bibliography section.");

        // All steps completed successfully.
    }
}
