using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for the sample files.
        string docxPath = "ContentControls.docx";
        string pdfPath = "ContentControls.pdf";

        // -------------------------------------------------
        // 1. Create a DOCX with an inline plain‑text SDT.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Introductory paragraph.
        builder.Writeln("Below is a plain‑text content control with a placeholder:");

        // Create the content control (inline plain‑text SDT).
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "SamplePlainText",
            Tag = "PlainTextTag",
            IsShowingPlaceholderText = true // Show placeholder when empty.
        };

        // The SDT must contain at least one run (even if empty).
        sdt.AppendChild(new Run(doc, string.Empty));

        // Insert the SDT into a new paragraph.
        builder.Writeln(); // Start a new paragraph.
        builder.CurrentParagraph?.AppendChild(sdt);
        builder.Writeln(); // Move to the next line.

        // Save the DOCX file.
        doc.Save(docxPath);

        // -------------------------------------------------
        // 2. Load the DOCX and convert it to PDF preserving placeholders.
        // -------------------------------------------------
        Document loadedDoc = new Document(docxPath);

        // Configure PDF save options to keep form fields and use the Tag as the field name.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            PreserveFormFields = true,
            UseSdtTagAsFormFieldName = true
        };

        // Save as PDF.
        loadedDoc.Save(pdfPath, pdfOptions);

        // Inform the user (no interactive input required).
        Console.WriteLine($"Conversion completed. PDF saved to '{Path.GetFullPath(pdfPath)}'.");
    }
}
