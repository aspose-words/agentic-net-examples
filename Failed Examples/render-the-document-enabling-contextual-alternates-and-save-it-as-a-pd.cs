// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source document (replace with your actual file path).
        Document doc = new Document("input.docx");

        // Enable contextual alternates to allow proper glyph substitution during rendering.
        doc.LayoutOptions.ContextualAlternates = true;

        // Rebuild the page layout to ensure the changes are taken into account.
        doc.UpdatePageLayout();

        // Save the rendered document as a PDF file.
        doc.Save("output.pdf");
    }
}
