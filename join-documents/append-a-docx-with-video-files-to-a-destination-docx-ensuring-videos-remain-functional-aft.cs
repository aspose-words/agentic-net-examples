using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file paths in the current directory.
        string destDocPath = Path.Combine(Directory.GetCurrentDirectory(), "Destination.docx");
        string srcDocPath = Path.Combine(Directory.GetCurrentDirectory(), "SourceWithVideo.docx");
        string mergedDocPath = Path.Combine(Directory.GetCurrentDirectory(), "Merged.docx");
        string pdfPath = Path.Combine(Directory.GetCurrentDirectory(), "Merged.pdf");

        // -----------------------------------------------------------------
        // 1. Create the destination document.
        // -----------------------------------------------------------------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("This is the destination document.");
        destDoc.Save(destDocPath);

        // -----------------------------------------------------------------
        // 2. Create the source document that contains an online video.
        // -----------------------------------------------------------------
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.Writeln("This document contains an online video.");

        // Insert an online video (YouTube example). The video will be embedded as a shape.
        // The URL must be a supported online video URL; YouTube works in Word.
        string videoUrl = "https://www.youtube.com/watch?v=dQw4w9WgXcQ";
        srcBuilder.InsertOnlineVideo(videoUrl, 320, 240);
        srcDoc.Save(srcDocPath);

        // -----------------------------------------------------------------
        // 3. Append the source document to the destination document.
        // -----------------------------------------------------------------
        // Reload the documents to ensure they are read from disk.
        Document destination = new Document(destDocPath);
        Document source = new Document(srcDocPath);

        // Append while keeping the source formatting (including the video shape).
        destination.AppendDocument(source, ImportFormatMode.KeepSourceFormatting);
        destination.Save(mergedDocPath);

        // -----------------------------------------------------------------
        // 4. Convert the merged document to PDF, embedding attachments (videos).
        // -----------------------------------------------------------------
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Embed video objects as annotations so they remain functional in the PDF.
            AttachmentsEmbeddingMode = PdfAttachmentsEmbeddingMode.Annotations
        };
        destination.Save(pdfPath, pdfOptions);

        // -----------------------------------------------------------------
        // 5. Validate that the output files were created.
        // -----------------------------------------------------------------
        if (!File.Exists(mergedDocPath))
            throw new InvalidOperationException($"Merged DOCX was not created at '{mergedDocPath}'.");
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException($"PDF was not created at '{pdfPath}'.");

        // Optional: Output the paths for verification (no interactive input required).
        Console.WriteLine($"Merged DOCX saved to: {mergedDocPath}");
        Console.WriteLine($"PDF with embedded video saved to: {pdfPath}");
    }
}
