using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Notes;

public class Program
{
    public static void Main()
    {
        // Paths for the output files
        string destinationPath = "Destination.docx";
        string sourcePath = "SourceWithFootnotes.docx";
        string mergedPdfPath = "MergedDocument.pdf";

        // -----------------------------------------------------------------
        // Create the destination document with a footnote
        // -----------------------------------------------------------------
        Document destinationDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destinationDoc);
        destBuilder.Writeln("This is the destination document.");
        destBuilder.InsertFootnote(FootnoteType.Footnote, "First footnote in destination.");
        // Ensure continuous footnote numbering
        destinationDoc.FootnoteOptions.RestartRule = FootnoteNumberingRule.Continuous;
        destinationDoc.Save(destinationPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Create the source document that also contains footnotes
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("This is the source document that will be appended.");
        srcBuilder.InsertFootnote(FootnoteType.Footnote, "First footnote in source.");
        srcBuilder.InsertFootnote(FootnoteType.Footnote, "Second footnote in source.");
        // Ensure continuous footnote numbering in the source as well
        sourceDoc.FootnoteOptions.RestartRule = FootnoteNumberingRule.Continuous;
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Load the documents (simulating a real‑world scenario)
        // -----------------------------------------------------------------
        Document dst = new Document(destinationPath);
        Document src = new Document(sourcePath);

        // Append the source document to the destination document.
        // Keep source formatting; footnote numbering will continue because both
        // documents use the Continuous restart rule.
        dst.AppendDocument(src, ImportFormatMode.KeepSourceFormatting);

        // Save the merged document as PDF.
        dst.Save(mergedPdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // Validation: ensure the PDF was created
        // -----------------------------------------------------------------
        if (!File.Exists(mergedPdfPath))
        {
            throw new InvalidOperationException($"Failed to create the merged PDF at '{mergedPdfPath}'.");
        }

        // Optional: clean up intermediate files (comment out if you need them)
        // File.Delete(destinationPath);
        // File.Delete(sourcePath);
    }
}
