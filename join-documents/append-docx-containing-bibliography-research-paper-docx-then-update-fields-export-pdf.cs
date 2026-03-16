using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Paths to the source documents and the output PDF.
        string researchPaperPath = "ResearchPaper.docx";
        string bibliographyPath = "Bibliography.docx";
        string outputPdfPath = "Combined.pdf";

        // Load the research paper document.
        Document researchPaper = new Document(researchPaperPath);

        // Load the bibliography document.
        Document bibliography = new Document(bibliographyPath);

        // Append the bibliography to the end of the research paper.
        // Keep the original formatting of the bibliography.
        researchPaper.AppendDocument(bibliography, ImportFormatMode.KeepSourceFormatting);

        // Update all fields in the combined document (e.g., TOC, citations, page numbers).
        researchPaper.UpdateFields();

        // Save the combined document as a PDF.
        researchPaper.Save(outputPdfPath, SaveFormat.Pdf);
    }
}
