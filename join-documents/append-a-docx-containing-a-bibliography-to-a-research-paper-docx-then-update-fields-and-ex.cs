using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for the sample documents and the final PDF.
        const string researchPath = "ResearchPaper.docx";
        const string bibliographyPath = "Bibliography.docx";
        const string outputPdfPath = "Combined.pdf";

        // -----------------------------------------------------------------
        // Create a sample research paper DOCX.
        // -----------------------------------------------------------------
        var researchDoc = new Document();
        var researchBuilder = new DocumentBuilder(researchDoc);
        researchBuilder.Writeln("Research Paper Title");
        researchBuilder.Writeln("This is the main content of the research paper.");
        researchBuilder.Writeln("It may contain multiple paragraphs, tables, images, etc.");
        researchDoc.Save(researchPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Create a sample bibliography DOCX.
        // -----------------------------------------------------------------
        var bibliographyDoc = new Document();
        var bibBuilder = new DocumentBuilder(bibliographyDoc);
        bibBuilder.Writeln("Bibliography");
        bibBuilder.Writeln("1. Author A. Title A. Publisher A, 2020.");
        bibBuilder.Writeln("2. Author B. Title B. Publisher B, 2021.");
        bibliographyDoc.Save(bibliographyPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Load the documents for joining.
        // -----------------------------------------------------------------
        var combinedDoc = new Document(researchPath);
        var bibToAppend = new Document(bibliographyPath);

        // Append the bibliography to the research paper.
        combinedDoc.AppendDocument(bibToAppend, ImportFormatMode.KeepSourceFormatting);

        // Update all fields (e.g., TOC, page numbers) in the combined document.
        combinedDoc.UpdateFields();

        // Save the combined document as PDF.
        combinedDoc.Save(outputPdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // Validation: ensure the PDF was created and contains content from both sources.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPdfPath))
        {
            throw new InvalidOperationException($"Failed to create the output PDF at '{outputPdfPath}'.");
        }

        // Simple validation: the combined document should have at least two sections
        // (one for each original document) unless they were merged into a single section.
        if (combinedDoc.Sections.Count < 2)
        {
            // If sections were merged, we can still check for presence of bibliography text.
            string pdfText = File.ReadAllText(outputPdfPath);
            if (!pdfText.Contains("Bibliography") || !pdfText.Contains("Research Paper Title"))
            {
                throw new InvalidOperationException("The combined document does not contain expected content from both source files.");
            }
        }

        // Indicate successful completion (optional, not required for non-interactive run).
        Console.WriteLine("Documents merged and PDF created successfully.");
    }
}
