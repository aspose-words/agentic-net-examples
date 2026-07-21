using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Notes;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // File paths.
        string destPath = Path.Combine(outputDir, "Destination.docx");
        string srcPath = Path.Combine(outputDir, "Source.docx");
        string mergedPath = Path.Combine(outputDir, "Merged.docx");
        string pdfPath = Path.Combine(outputDir, "Merged.pdf");

        // ---------- Create destination document with footnotes ----------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("Destination document start.");
        destBuilder.InsertFootnote(FootnoteType.Footnote, "Destination footnote 1.");
        destBuilder.Writeln("More text in destination.");
        destBuilder.InsertFootnote(FootnoteType.Footnote, "Destination footnote 2.");
        destDoc.Save(destPath);

        // ---------- Create source document with footnotes ----------
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.Writeln("Source document start.");
        srcBuilder.InsertFootnote(FootnoteType.Footnote, "Source footnote 1.");
        srcBuilder.Writeln("More text in source.");
        srcBuilder.InsertFootnote(FootnoteType.Footnote, "Source footnote 2.");
        srcDoc.Save(srcPath);

        // ---------- Append source to destination ----------
        // Load the previously saved documents (optional, we could use the objects directly).
        Document destination = new Document(destPath);
        Document source = new Document(srcPath);

        // Append while keeping source formatting.
        destination.AppendDocument(source, ImportFormatMode.KeepSourceFormatting);

        // Ensure footnote numbering continues throughout the merged document.
        destination.FootnoteOptions.RestartRule = FootnoteNumberingRule.Continuous;

        // Save the merged document.
        destination.Save(mergedPath);

        // Save the merged document as PDF.
        destination.Save(pdfPath, SaveFormat.Pdf);

        // ---------- Validation ----------
        if (!File.Exists(pdfPath))
            throw new Exception("PDF file was not created.");

        // Optional: output paths for verification (no console interaction required).
        Console.WriteLine("Merged DOCX: " + mergedPath);
        Console.WriteLine("Merged PDF: " + pdfPath);
    }
}
