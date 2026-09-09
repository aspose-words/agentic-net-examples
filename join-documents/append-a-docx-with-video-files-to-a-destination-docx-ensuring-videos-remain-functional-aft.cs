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

        // -----------------------------------------------------------------
        // 1. Create a dummy video file (placeholder content).
        // -----------------------------------------------------------------
        string videoPath = Path.Combine(outputDir, "sample.mp4");
        // Write a few bytes to make the file exist; real video content is not required for the demo.
        File.WriteAllBytes(videoPath, new byte[] { 0x00, 0x01, 0x02, 0x03 });

        // -----------------------------------------------------------------
        // 2. Create the source DOCX that contains the video.
        // -----------------------------------------------------------------
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.Writeln("Source document with an embedded video:");
        // Embed the video as an OLE object. Use the overload that accepts (fileName, isLinked, asIcon, presentation).
        srcBuilder.InsertOleObject(videoPath, isLinked: false, asIcon: false, presentation: null);
        string srcPath = Path.Combine(outputDir, "Source.docx");
        srcDoc.Save(srcPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 3. Create the destination DOCX.
        // -----------------------------------------------------------------
        Document dstDoc = new Document();
        DocumentBuilder dstBuilder = new DocumentBuilder(dstDoc);
        dstBuilder.Writeln("Destination document (will receive the source).");
        string dstPath = Path.Combine(outputDir, "Destination.docx");
        dstDoc.Save(dstPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 4. Append the source document to the destination document.
        // -----------------------------------------------------------------
        // Load the documents again to simulate a real‑world scenario.
        Document destination = new Document(dstPath);
        Document source = new Document(srcPath);
        destination.AppendDocument(source, ImportFormatMode.KeepSourceFormatting);
        string mergedPath = Path.Combine(outputDir, "Merged.docx");
        destination.Save(mergedPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 5. Convert the merged document to PDF, embedding the video attachment.
        // -----------------------------------------------------------------
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            AttachmentsEmbeddingMode = PdfAttachmentsEmbeddingMode.Annotations
        };
        string pdfPath = Path.Combine(outputDir, "Merged.pdf");
        destination.Save(pdfPath, pdfOptions);

        // -----------------------------------------------------------------
        // 6. Validate that all output files were created.
        // -----------------------------------------------------------------
        ValidateFileExists(srcPath);
        ValidateFileExists(dstPath);
        ValidateFileExists(mergedPath);
        ValidateFileExists(pdfPath);
        ValidateFileExists(videoPath);
    }

    private static void ValidateFileExists(string path)
    {
        if (!File.Exists(path))
        {
            throw new FileNotFoundException($"Expected file was not created: {path}");
        }
    }
}
