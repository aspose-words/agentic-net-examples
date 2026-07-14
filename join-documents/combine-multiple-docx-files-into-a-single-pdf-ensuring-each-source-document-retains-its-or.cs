using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // File paths for the sample DOCX files and the final PDF
        string doc1Path = Path.Combine(outputDir, "Doc1.docx");
        string doc2Path = Path.Combine(outputDir, "Doc2.docx");
        string mergedPdfPath = Path.Combine(outputDir, "Merged.pdf");

        // ---------- Create first sample DOCX ----------
        Document doc1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(doc1);
        builder1.Writeln("First Document - Heading");
        builder1.Font.Size = 14;
        builder1.Writeln("This is the first sample document.");
        doc1.Save(doc1Path, SaveFormat.Docx);

        // ---------- Create second sample DOCX ----------
        Document doc2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(doc2);
        builder2.Writeln("Second Document - Heading");
        builder2.Font.Size = 14;
        builder2.Font.Color = Color.Blue;
        builder2.Writeln("This is the second sample document with a different style.");
        doc2.Save(doc2Path, SaveFormat.Docx);

        // ---------- Load the source documents ----------
        Document srcDoc1 = new Document(doc1Path);
        Document srcDoc2 = new Document(doc2Path);

        // ---------- Destination document ----------
        Document dstDoc = new Document();

        // Append the first source document while preserving its formatting
        dstDoc.AppendDocument(srcDoc1, ImportFormatMode.KeepSourceFormatting);

        // Insert a page break between the merged documents
        DocumentBuilder dstBuilder = new DocumentBuilder(dstDoc);
        dstBuilder.MoveToDocumentEnd();
        dstBuilder.InsertBreak(BreakType.PageBreak);

        // Append the second source document while preserving its formatting
        dstDoc.AppendDocument(srcDoc2, ImportFormatMode.KeepSourceFormatting);

        // ---------- Save the combined document as PDF ----------
        dstDoc.Save(mergedPdfPath, SaveFormat.Pdf);

        // ---------- Validate that the PDF was created ----------
        if (!File.Exists(mergedPdfPath))
        {
            throw new InvalidOperationException("Merged PDF was not created.");
        }

        // Simple confirmation (non‑interactive)
        Console.WriteLine($"Merged PDF created at: {mergedPdfPath}");
    }
}
