using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        string videoPath = "sample_video.mp4";
        string destinationDocPath = "Destination.docx";
        string sourceDocPath = "Source.docx";
        string mergedPdfPath = "MergedOutput.pdf";

        // Create a dummy video file (placeholder content)
        File.WriteAllBytes(videoPath, new byte[] { 0x00, 0x01, 0x02, 0x03 });

        // -------------------------------------------------
        // Create the destination DOCX with a video embedded as an OLE object
        // -------------------------------------------------
        var destinationDoc = new Document();
        var destBuilder = new DocumentBuilder(destinationDoc);
        destBuilder.Writeln("This is the destination document.");

        using (FileStream videoStream = File.OpenRead(videoPath))
        {
            // Insert the video as a generic package OLE object.
            // The fourth parameter is the icon stream; we pass null because we do not provide a custom icon.
            destBuilder.InsertOleObject(videoStream, "Package", false, (Stream)null);
        }

        destinationDoc.Save(destinationDocPath, SaveFormat.Docx);

        // -------------------------------------------------
        // Create the source DOCX with a video embedded as an OLE object
        // -------------------------------------------------
        var sourceDoc = new Document();
        var srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("This is the source document.");

        using (FileStream videoStream = File.OpenRead(videoPath))
        {
            srcBuilder.InsertOleObject(videoStream, "Package", false, (Stream)null);
        }

        sourceDoc.Save(sourceDocPath, SaveFormat.Docx);

        // -------------------------------------------------
        // Load the documents from disk
        // -------------------------------------------------
        var destDocToMerge = new Document(destinationDocPath);
        var srcDocToMerge = new Document(sourceDocPath);

        // Append the source document to the destination document, preserving source formatting
        destDocToMerge.AppendDocument(srcDocToMerge, ImportFormatMode.KeepSourceFormatting);

        // Save the merged document as PDF; embedded videos are kept as OLE objects in the PDF
        destDocToMerge.Save(mergedPdfPath, SaveFormat.Pdf);

        // Validate that the PDF was created
        if (!File.Exists(mergedPdfPath))
        {
            throw new InvalidOperationException("Merged PDF was not created.");
        }

        // Clean up temporary files
        File.Delete(videoPath);
        File.Delete(destinationDocPath);
        File.Delete(sourceDocPath);
    }
}
