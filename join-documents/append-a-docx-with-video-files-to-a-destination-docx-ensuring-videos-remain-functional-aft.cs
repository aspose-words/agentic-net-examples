using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a dummy video file (the content is not important for the example).
        string videoPath = Path.Combine(outputDir, "sample.mp4");
        File.WriteAllBytes(videoPath, new byte[] { 0x00, 0x01, 0x02, 0x03 });

        // -------------------- Create source DOCX with an embedded video --------------------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("Source document with video:");
        // Embed the video file as an OLE object (embedded, not linked, not shown as an icon).
        srcBuilder.InsertOleObject(videoPath, isLinked: false, asIcon: false, presentation: null);
        string sourceDocPath = Path.Combine(outputDir, "Source.docx");
        sourceDoc.Save(sourceDocPath, SaveFormat.Docx);

        // -------------------- Create destination DOCX --------------------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("Destination document header.");
        string destDocPath = Path.Combine(outputDir, "Destination.docx");
        destDoc.Save(destDocPath, SaveFormat.Docx);

        // -------------------- Append source document to destination --------------------
        Document srcToAppend = new Document(sourceDocPath);
        destDoc.AppendDocument(srcToAppend, ImportFormatMode.KeepSourceFormatting);
        string mergedDocPath = Path.Combine(outputDir, "Merged.docx");
        destDoc.Save(mergedDocPath, SaveFormat.Docx);

        // -------------------- Save merged document as PDF with embedded video --------------------
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Ensure OLE objects (the video) are embedded in the PDF as annotations.
            AttachmentsEmbeddingMode = PdfAttachmentsEmbeddingMode.Annotations
        };
        string pdfPath = Path.Combine(outputDir, "Merged.pdf");
        destDoc.Save(pdfPath, pdfOptions);

        // -------------------- Validation --------------------
        if (!File.Exists(mergedDocPath) || !File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the expected output files.");

        Console.WriteLine($"Merged DOCX created at: {mergedDocPath}");
        Console.WriteLine($"Merged PDF created at: {pdfPath}");
    }
}
