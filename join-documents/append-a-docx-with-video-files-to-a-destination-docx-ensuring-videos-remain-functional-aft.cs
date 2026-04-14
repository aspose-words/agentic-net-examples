using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file paths in the current directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        string destDocPath = Path.Combine(outputDir, "Destination.docx");
        string srcDocPath = Path.Combine(outputDir, "SourceWithVideo.docx");
        string mergedDocPath = Path.Combine(outputDir, "Merged.docx");
        string mergedPdfPath = Path.Combine(outputDir, "Merged.pdf");
        string videoPath = Path.Combine(outputDir, "sample.mp4");

        // -----------------------------------------------------------------
        // 1. Create a dummy video file (the content is not important for the demo).
        // -----------------------------------------------------------------
        byte[] dummyVideoContent = new byte[] { 0x00, 0x01, 0x02, 0x03, 0x04 };
        File.WriteAllBytes(videoPath, dummyVideoContent);

        // -----------------------------------------------------------------
        // 2. Create the destination document.
        // -----------------------------------------------------------------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("This is the destination document.");
        destDoc.Save(destDocPath);

        // -----------------------------------------------------------------
        // 3. Create the source document that contains a video (embedded as OLE object).
        // -----------------------------------------------------------------
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.Writeln("This is the source document with an embedded video.");

        // Embed the video file as an OLE object.
        // Parameters: file name, isLink (false = embed), updateFields (true), imageData (null = default icon).
        srcBuilder.InsertOleObject(videoPath, false, true, null);
        srcDoc.Save(srcDocPath);

        // -----------------------------------------------------------------
        // 4. Append the source document to the destination document.
        // -----------------------------------------------------------------
        // Load the documents (they are already in memory, but loading from file ensures a realistic scenario).
        Document destination = new Document(destDocPath);
        Document source = new Document(srcDocPath);

        // Append while keeping the source formatting.
        destination.AppendDocument(source, ImportFormatMode.KeepSourceFormatting);
        destination.Save(mergedDocPath);

        // -----------------------------------------------------------------
        // 5. Save the merged document as PDF with embedded attachments (videos).
        // -----------------------------------------------------------------
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            AttachmentsEmbeddingMode = PdfAttachmentsEmbeddingMode.Annotations
        };
        destination.Save(mergedPdfPath, pdfOptions);

        // -----------------------------------------------------------------
        // 6. Validate that the output files were created.
        // -----------------------------------------------------------------
        if (!File.Exists(mergedDocPath))
            throw new FileNotFoundException("Merged DOCX was not created.", mergedDocPath);
        if (!File.Exists(mergedPdfPath))
            throw new FileNotFoundException("Merged PDF was not created.", mergedPdfPath);

        // The example finishes here. No interactive prompts are used.
    }
}
