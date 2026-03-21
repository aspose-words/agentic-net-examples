using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Paths to the source documents and the output PDF.
        const string researchPaperPath = "ResearchPaper.docx";
        const string bibliographyPath = "Bibliography.docx";
        const string outputPdfPath = "Combined.pdf";

        // Ensure the source documents exist. If they don't, create simple placeholder documents.
        if (!System.IO.File.Exists(researchPaperPath))
        {
            var researchDoc = new Document();
            var builder = new DocumentBuilder(researchDoc);
            builder.Writeln("Research Paper Title");
            builder.Writeln("This is the content of the research paper.");
            researchDoc.Save(researchPaperPath);
        }

        if (!System.IO.File.Exists(bibliographyPath))
        {
            var bibliographyDoc = new Document();
            var builder = new DocumentBuilder(bibliographyDoc);
            builder.Writeln("Bibliography");
            builder.Writeln("1. Author A. Title A. Journal A, 2020.");
            builder.Writeln("2. Author B. Title B. Journal B, 2021.");
            bibliographyDoc.Save(bibliographyPath);
        }

        // Load the research paper document.
        Document researchPaper = new Document(researchPaperPath);

        // Load the bibliography document.
        Document bibliography = new Document(bibliographyPath);

        // Append the bibliography to the end of the research paper, keeping its original formatting.
        researchPaper.AppendDocument(bibliography, ImportFormatMode.KeepSourceFormatting);

        // Update all fields in the combined document (e.g., TOC, citations, page numbers).
        researchPaper.UpdateFields();

        // Save the combined document as a PDF.
        researchPaper.Save(outputPdfPath, SaveFormat.Pdf);

        Console.WriteLine($"Combined PDF saved to '{outputPdfPath}'.");
    }
}
