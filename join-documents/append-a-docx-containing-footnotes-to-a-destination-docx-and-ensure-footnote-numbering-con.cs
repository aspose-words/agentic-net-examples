using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Notes;

public class Program
{
    public static void Main()
    {
        // Create the destination document.
        Document destination = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destination);
        destBuilder.Writeln("Destination document content.");

        // Create the source document that contains footnotes.
        Document source = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(source);
        srcBuilder.Writeln("Source document content with footnotes.");
        srcBuilder.Write("Reference to first footnote. ");
        srcBuilder.InsertFootnote(FootnoteType.Footnote, "First footnote text.");
        srcBuilder.Writeln();
        srcBuilder.Write("Reference to second footnote. ");
        srcBuilder.InsertFootnote(FootnoteType.Footnote, "Second footnote text.");
        srcBuilder.Writeln();

        // Append the source document to the destination.
        // KeepSourceFormatting preserves the original formatting.
        destination.AppendDocument(source, ImportFormatMode.KeepSourceFormatting);

        // Save the merged document as PDF.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedDocument.pdf");
        destination.Save(outputPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The merged PDF file was not created.");
    }
}
