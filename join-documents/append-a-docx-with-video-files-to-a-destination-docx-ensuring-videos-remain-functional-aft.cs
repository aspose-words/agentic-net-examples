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

        // Define file paths.
        string destDocPath = Path.Combine(outputDir, "Destination.docx");
        string sourceDocPath = Path.Combine(outputDir, "SourceWithVideo.docx");
        string mergedDocPath = Path.Combine(outputDir, "Merged.docx");
        string mergedPdfPath = Path.Combine(outputDir, "Merged.pdf");
        string videoFilePath = Path.Combine(outputDir, "sample.mp4");

        // Create a dummy video file (binary content).
        File.WriteAllBytes(videoFilePath, new byte[] { 0x00, 0x01, 0x02, 0x03, 0x04 });

        // ---------- Create Destination Document ----------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("This is the destination document.");
        destDoc.Save(destDocPath, SaveFormat.Docx);

        // ---------- Create Source Document with Embedded Video ----------
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);
        sourceBuilder.Writeln("This is the source document containing a video.");

        // Insert the video as an OLE object (embedded) using a stream.
        using (FileStream videoStream = File.OpenRead(videoFilePath))
        {
            // The fourth parameter (icon stream) is optional; pass null for no icon.
            sourceBuilder.InsertOleObject(videoStream, "Package", false, null);
        }

        sourceDoc.Save(sourceDocPath, SaveFormat.Docx);

        // ---------- Load Documents ----------
        Document destination = new Document(destDocPath);
        Document source = new Document(sourceDocPath);

        // ---------- Append Source to Destination ----------
        destination.AppendDocument(source, ImportFormatMode.KeepSourceFormatting);
        destination.Save(mergedDocPath, SaveFormat.Docx);

        // ---------- Convert Merged Document to PDF ----------
        destination.Save(mergedPdfPath, SaveFormat.Pdf);

        // ---------- Validation ----------
        if (!File.Exists(mergedDocPath))
            throw new Exception("Merged DOCX file was not created.");

        if (!File.Exists(mergedPdfPath))
            throw new Exception("Merged PDF file was not created.");

        if (new FileInfo(mergedPdfPath).Length == 0)
            throw new Exception("Merged PDF file is empty.");
    }
}
