using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file paths for the sample documents and the final PDF.
        string researchPath = "ResearchPaper.docx";
        string bibliographyPath = "Bibliography.docx";
        string outputPdfPath = "ResearchPaper_WithBibliography.pdf";

        // -----------------------------------------------------------------
        // Create a sample research paper document.
        // -----------------------------------------------------------------
        Document researchDoc = new Document();
        DocumentBuilder researchBuilder = new DocumentBuilder(researchDoc);
        researchBuilder.Writeln("Research Paper Title");
        researchBuilder.Writeln("This is the content of the research paper.");
        // Insert a simple field to demonstrate field updating.
        researchBuilder.InsertField("PAGE", null);
        researchBuilder.Writeln(); // Ensure a new paragraph after the field.
        researchDoc.Save(researchPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Create a sample bibliography document.
        // -----------------------------------------------------------------
        Document bibliographyDoc = new Document();
        DocumentBuilder bibBuilder = new DocumentBuilder(bibliographyDoc);
        bibBuilder.Writeln("Bibliography");
        bibBuilder.Writeln("1. Reference A");
        bibBuilder.Writeln("2. Reference B");
        bibliographyDoc.Save(bibliographyPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Load the created documents (optional, as they are already in memory).
        // -----------------------------------------------------------------
        Document research = new Document(researchPath);
        Document bibliography = new Document(bibliographyPath);

        // Append the bibliography to the research paper, preserving source formatting.
        research.AppendDocument(bibliography, ImportFormatMode.KeepSourceFormatting);

        // Update all fields in the combined document (e.g., PAGE field).
        research.UpdateFields();

        // Save the combined document as PDF.
        research.Save(outputPdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // Validation: ensure the PDF was created and contains bibliography text.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPdfPath))
        {
            throw new InvalidOperationException($"The output PDF '{outputPdfPath}' was not created.");
        }

        // Load the PDF back as a Document to verify its content.
        Document resultDoc = new Document(outputPdfPath);
        string resultText = resultDoc.GetText();

        if (!resultText.Contains("Bibliography") || !resultText.Contains("Reference A"))
        {
            throw new InvalidOperationException("The merged PDF does not contain expected bibliography content.");
        }

        // Optionally, clean up the intermediate files (comment out if inspection is needed).
        // File.Delete(researchPath);
        // File.Delete(bibliographyPath);
    }
}
