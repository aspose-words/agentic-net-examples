using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the sample documents.
        string sourceDocPath = Path.Combine(outputDir, "Source.docx");
        string destinationDocPath = Path.Combine(outputDir, "Destination.docx");
        string mergedDocPath = Path.Combine(outputDir, "Merged.docx");
        string mergedPdfPath = Path.Combine(outputDir, "Merged.pdf");

        // ---------- Create source DOCX containing an online video ----------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("Source document with an embedded online video:");
        // Insert an online video (YouTube or any public video URL). Size is in points.
        srcBuilder.InsertOnlineVideo(
            "https://sample-videos.com/video123/mp4/720/big_buck_bunny_720p_1mb.mp4",
            400,
            300);
        sourceDoc.Save(sourceDocPath, SaveFormat.Docx);

        // ---------- Create destination DOCX ----------
        Document destinationDoc = new Document();
        DocumentBuilder dstBuilder = new DocumentBuilder(destinationDoc);
        dstBuilder.Writeln("Destination document – introductory text.");
        destinationDoc.Save(destinationDocPath, SaveFormat.Docx);

        // ---------- Append source document to destination document ----------
        // Load the documents (already in memory, but loading from file demonstrates the workflow).
        Document dstDoc = new Document(destinationDocPath);
        Document srcDoc = new Document(sourceDocPath);
        dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
        dstDoc.Save(mergedDocPath, SaveFormat.Docx);

        // ---------- Convert the merged document to PDF, preserving the video ----------
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Embed video as an annotation so it remains functional in the PDF.
            AttachmentsEmbeddingMode = PdfAttachmentsEmbeddingMode.Annotations
        };
        dstDoc.Save(mergedPdfPath, pdfOptions);

        // ---------- Validation ----------
        if (!File.Exists(mergedPdfPath))
        {
            throw new InvalidOperationException("The merged PDF was not created.");
        }

        // The example finishes without requiring user interaction.
    }
}
