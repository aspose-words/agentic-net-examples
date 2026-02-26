// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load an existing Word document.
        Document doc = new Document("Input.docx");

        // Enable contextual alternates for proper glyph substitution.
        doc.LayoutOptions.ContextualAlternates = true;

        // Ensure the page layout is up‑to‑date before rendering.
        doc.UpdatePageLayout();

        // Create PDF save options (default settings are sufficient for this task).
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the document as PDF, applying the enabled contextual alternates.
        doc.Save("Output.pdf", pdfOptions);
    }
}
