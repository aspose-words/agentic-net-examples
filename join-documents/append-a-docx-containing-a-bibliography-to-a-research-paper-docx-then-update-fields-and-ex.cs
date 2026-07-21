using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Bibliography;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for temporary files.
        string dataDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(dataDir);

        // Paths for the source and output documents.
        string researchPath = Path.Combine(dataDir, "ResearchPaper.docx");
        string bibliographyPath = Path.Combine(dataDir, "Bibliography.docx");
        string mergedPath = Path.Combine(dataDir, "Merged.docx");
        string pdfPath = Path.Combine(dataDir, "Merged.pdf");

        // -------------------------
        // Create the research paper.
        // -------------------------
        Document researchDoc = new Document();
        DocumentBuilder researchBuilder = new DocumentBuilder(researchDoc);
        researchBuilder.Writeln("Research Paper Title");
        researchBuilder.Writeln("This is the content of the research paper.");
        // Insert a simple PAGE field to demonstrate field updating.
        researchBuilder.InsertField(FieldType.FieldPage, true);
        researchDoc.Save(researchPath);

        // -------------------------
        // Create the bibliography document.
        // -------------------------
        Document bibDoc = new Document();
        DocumentBuilder bibBuilder = new DocumentBuilder(bibDoc);
        // Add a sample source to the bibliography collection.
        bibDoc.Bibliography.Sources.Add(new Source("Sample Book", SourceType.Book));
        // Write a heading and insert a BIBLIOGRAPHY field.
        bibBuilder.Writeln("Bibliography:");
        bibBuilder.InsertField(FieldType.FieldBibliography, true);
        bibDoc.Save(bibliographyPath);

        // -------------------------
        // Load the documents and join them.
        // -------------------------
        Document research = new Document(researchPath);
        Document bibliography = new Document(bibliographyPath);

        // Append the bibliography document to the research paper,
        // preserving the source formatting.
        research.AppendDocument(bibliography, ImportFormatMode.KeepSourceFormatting);

        // Update all fields (including the PAGE and BIBLIOGRAPHY fields).
        research.UpdateFields();

        // Save the merged document in DOCX format.
        research.Save(mergedPath);

        // Export the merged document to PDF.
        research.Save(pdfPath, SaveFormat.Pdf);

        // -------------------------
        // Validation.
        // -------------------------
        if (!File.Exists(mergedPath))
            throw new Exception("Merged DOCX file was not created.");

        if (!File.Exists(pdfPath))
            throw new Exception("PDF file was not created.");

        string mergedText = research.GetText();

        if (!mergedText.Contains("Research Paper Title"))
            throw new Exception("Merged document does not contain research paper content.");

        if (!mergedText.Contains("Bibliography"))
            throw new Exception("Merged document does not contain bibliography content.");

        // All operations completed successfully.
    }
}
