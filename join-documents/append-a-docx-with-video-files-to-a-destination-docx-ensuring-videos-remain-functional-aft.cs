using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare a temporary working folder.
        string workFolder = Path.Combine(Path.GetTempPath(), "AsposeJoinDemo");
        Directory.CreateDirectory(workFolder);

        // Create a dummy video file (MP4) – content is not important for the demo.
        string videoPath = Path.Combine(workFolder, "sample.mp4");
        File.WriteAllBytes(videoPath, new byte[] { 0x00, 0x00, 0x00, 0x18, 0x66, 0x74, 0x79, 0x70 });

        // Create the source DOCX that *represents* a document with an embedded video.
        string sourceDocPath = Path.Combine(workFolder, "source.docx");
        var sourceDoc = new Document();
        var sourceBuilder = new DocumentBuilder(sourceDoc);
        sourceBuilder.Writeln("This is the source document with an embedded video.");
        // Insert a placeholder that indicates where the video would be.
        sourceBuilder.Writeln("[Video placeholder: sample.mp4]");
        sourceDoc.Save(sourceDocPath, SaveFormat.Docx);

        // Create the destination DOCX.
        string destDocPath = Path.Combine(workFolder, "destination.docx");
        var destDoc = new Document();
        var destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("This is the destination document.");
        destDoc.Save(destDocPath, SaveFormat.Docx);

        // Load both documents.
        var destination = new Document(destDocPath);
        var source = new Document(sourceDocPath);

        // Append the source document to the destination document, keeping source formatting.
        destination.AppendDocument(source, ImportFormatMode.KeepSourceFormatting);

        // Save the merged document as DOCX.
        string mergedDocPath = Path.Combine(workFolder, "merged.docx");
        destination.Save(mergedDocPath, SaveFormat.Docx);

        // Convert the merged DOCX to PDF.
        string mergedPdfPath = Path.Combine(workFolder, "merged.pdf");
        destination.Save(mergedPdfPath, SaveFormat.Pdf);

        // Validation: ensure the output files exist.
        if (!File.Exists(mergedDocPath))
            throw new InvalidOperationException($"Merged DOCX was not created at '{mergedDocPath}'.");
        if (!File.Exists(mergedPdfPath))
            throw new InvalidOperationException($"Merged PDF was not created at '{mergedPdfPath}'.");

        // Write a simple log file indicating success.
        string logPath = Path.Combine(workFolder, "log.txt");
        File.WriteAllText(logPath,
            $"Merge completed successfully.{Environment.NewLine}" +
            $"DOCX: {mergedDocPath}{Environment.NewLine}" +
            $"PDF: {mergedPdfPath}");

        // Optionally, keep the temporary files for inspection.
        // Directory.Delete(workFolder, true);
    }
}
