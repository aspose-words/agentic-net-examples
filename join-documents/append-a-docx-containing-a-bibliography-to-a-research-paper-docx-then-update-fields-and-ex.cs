using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fields;   // Needed for FieldType enum

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        const string researchPath = "ResearchPaper.docx";
        const string bibliographyPath = "Bibliography.docx";
        const string mergedPath = "MergedResearchPaper.docx";
        const string pdfPath = "MergedResearchPaper.pdf";

        // -----------------------------------------------------------------
        // Create a sample research paper document.
        // -----------------------------------------------------------------
        Document researchDoc = new Document();
        DocumentBuilder researchBuilder = new DocumentBuilder(researchDoc);
        researchBuilder.Writeln("Research Paper Title");
        researchBuilder.Writeln("This is the introduction of the research paper.");
        researchBuilder.Writeln("Content goes here...");
        researchDoc.Save(researchPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Create a sample bibliography document containing a bibliography field.
        // -----------------------------------------------------------------
        Document bibliographyDoc = new Document();
        DocumentBuilder bibBuilder = new DocumentBuilder(bibliographyDoc);
        bibBuilder.Writeln("Bibliography");
        // Insert a BIBLIOGRAPHY field; it will be updated later.
        bibBuilder.InsertField(FieldType.FieldBibliography, true);
        bibliographyDoc.Save(bibliographyPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Load the two documents.
        // -----------------------------------------------------------------
        Document mainDoc = new Document(researchPath);
        Document srcBibliography = new Document(bibliographyPath);

        // -----------------------------------------------------------------
        // Append the bibliography document to the research paper.
        // Keep the source formatting.
        // -----------------------------------------------------------------
        mainDoc.AppendDocument(srcBibliography, ImportFormatMode.KeepSourceFormatting);

        // -----------------------------------------------------------------
        // Update all fields (including the bibliography field) in the merged document.
        // -----------------------------------------------------------------
        mainDoc.UpdateFields();

        // -----------------------------------------------------------------
        // Save the merged document as DOCX.
        // -----------------------------------------------------------------
        mainDoc.Save(mergedPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Export the merged document to PDF.
        // -----------------------------------------------------------------
        mainDoc.Save(pdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // Validate that the output files were created.
        // -----------------------------------------------------------------
        if (!File.Exists(mergedPath))
            throw new InvalidOperationException($"Merged DOCX file was not created: {mergedPath}");

        if (!File.Exists(pdfPath))
            throw new InvalidOperationException($"PDF file was not created: {pdfPath}");

        // Program ends without requiring any user interaction.
    }
}
