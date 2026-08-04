using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define a folder for all generated files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Paths for the dummy video, source DOCX, destination DOCX, merged DOCX and final PDF.
        string videoPath = Path.Combine(artifactsDir, "sample.mp4");
        string sourceDocPath = Path.Combine(artifactsDir, "source.docx");
        string destDocPath = Path.Combine(artifactsDir, "destination.docx");
        string mergedDocPath = Path.Combine(artifactsDir, "merged.docx");
        string mergedPdfPath = Path.Combine(artifactsDir, "merged.pdf");

        // Create a placeholder video file (the content is irrelevant for the demo).
        File.WriteAllBytes(videoPath, new byte[] { 0x00, 0x01, 0x02, 0x03 });

        // -------------------- Create source document with an embedded video --------------------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("Source document with an embedded video:");
        // Embed the video as an OLE object (not a link, not an icon).
        srcBuilder.InsertOleObject(videoPath, false, false, null);
        sourceDoc.Save(sourceDocPath);

        // -------------------- Create destination document --------------------
        Document destDoc = new Document();
        DocumentBuilder dstBuilder = new DocumentBuilder(destDoc);
        dstBuilder.Writeln("Destination document content.");
        destDoc.Save(destDocPath);

        // -------------------- Load documents and append --------------------
        Document src = new Document(sourceDocPath);
        Document dst = new Document(destDocPath);
        // Append the source document while preserving its formatting (including the OLE video).
        dst.AppendDocument(src, ImportFormatMode.KeepSourceFormatting);
        dst.Save(mergedDocPath);

        // -------------------- Convert merged document to PDF with video embedded --------------------
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Embed OLE objects (the video) as annotations so they remain functional in the PDF.
            AttachmentsEmbeddingMode = PdfAttachmentsEmbeddingMode.Annotations
        };
        dst.Save(mergedPdfPath, pdfOptions);

        // -------------------- Validation --------------------
        if (!File.Exists(mergedDocPath))
            throw new InvalidOperationException("Merged DOCX was not created.");

        if (!File.Exists(mergedPdfPath))
            throw new InvalidOperationException("Merged PDF was not created.");

        // Program ends without waiting for user input.
    }
}
