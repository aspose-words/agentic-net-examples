using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Notes;

public class Program
{
    public static void Main()
    {
        // Create the destination document with a footnote.
        Document destinationDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destinationDoc);
        destBuilder.Writeln("Destination document content.");
        destBuilder.InsertFootnote(FootnoteType.Footnote, "First footnote in destination.");

        // Ensure footnote numbering is continuous (default, but set explicitly).
        destinationDoc.FootnoteOptions.RestartRule = FootnoteNumberingRule.Continuous;

        // Create the source document that also contains footnotes.
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("Source document content.");
        srcBuilder.InsertFootnote(FootnoteType.Footnote, "First footnote in source.");
        srcBuilder.InsertFootnote(FootnoteType.Footnote, "Second footnote in source.");

        // Append the source document to the destination document.
        destinationDoc.AppendDocument(sourceDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the merged document as PDF.
        const string outputPdfPath = "MergedDocument.pdf";
        destinationDoc.Save(outputPdfPath, SaveFormat.Pdf);

        // Validate that the PDF file was created.
        if (!File.Exists(outputPdfPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPdfPath}");
    }
}
