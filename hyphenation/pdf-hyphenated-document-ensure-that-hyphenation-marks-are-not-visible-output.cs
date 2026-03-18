using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Layout;

class HyphenationToPdf
{
    static void Main()
    {
        // Load the source document that contains hyphenated text.
        // Replace the path with the actual location of your .docx file.
        Document doc = new Document("HyphenatedDocument.docx");

        // Enable automatic hyphenation so the layout engine can break words.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Suppress visible hyphenation marks for every paragraph.
        // This keeps the layout hyphenation but removes the hyphen characters from the output.
        foreach (Paragraph paragraph in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            paragraph.ParagraphFormat.SuppressAutoHyphens = true;
        }

        // Re‑layout the document to apply the changes.
        doc.UpdatePageLayout();

        // Prepare PDF save options (optional – can be omitted if defaults are sufficient).
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Ensure the document structure is exported (helps retain formatting).
            ExportDocumentStructure = true
        };

        // Save the result as a PDF. Replace the path with the desired output location.
        doc.Save("HyphenatedOutput.pdf", pdfOptions);
    }
}
