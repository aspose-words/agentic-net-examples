using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Notes;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // File names.
        string destPath = Path.Combine(outputDir, "Destination.docx");
        string srcPath = Path.Combine(outputDir, "Source.docx");
        string mergedPdfPath = Path.Combine(outputDir, "Merged.pdf");

        // ---------- Create destination document with footnotes ----------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("Destination Document");
        destBuilder.Write("First paragraph with a footnote. ");
        destBuilder.InsertFootnote(FootnoteType.Footnote, "Destination footnote 1.");
        destBuilder.Writeln();
        destBuilder.Write("Second paragraph with another footnote. ");
        destBuilder.InsertFootnote(FootnoteType.Footnote, "Destination footnote 2.");
        destDoc.Save(destPath);

        // ---------- Create source document with footnotes ----------
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.Writeln("Source Document");
        srcBuilder.Write("Paragraph in source with a footnote. ");
        srcBuilder.InsertFootnote(FootnoteType.Footnote, "Source footnote 1.");
        srcBuilder.Writeln();
        srcBuilder.Write("Another source paragraph with footnote. ");
        srcBuilder.InsertFootnote(FootnoteType.Footnote, "Source footnote 2.");
        srcDoc.Save(srcPath);

        // ---------- Load documents ----------
        Document destination = new Document(destPath);
        Document source = new Document(srcPath);

        // Ensure footnote numbering is continuous throughout the merged document.
        destination.FootnoteOptions.RestartRule = FootnoteNumberingRule.Continuous;

        // Append source document to destination, preserving source formatting.
        destination.AppendDocument(source, ImportFormatMode.KeepSourceFormatting);

        // Save the merged document as PDF.
        destination.Save(mergedPdfPath, SaveFormat.Pdf);

        // ---------- Validation ----------
        if (!File.Exists(destPath))
            throw new FileNotFoundException("Destination DOCX was not created.", destPath);
        if (!File.Exists(srcPath))
            throw new FileNotFoundException("Source DOCX was not created.", srcPath);
        if (!File.Exists(mergedPdfPath))
            throw new FileNotFoundException("Merged PDF was not created.", mergedPdfPath);
    }
}
