using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Simulate a network share by using a temporary folder.
        string networkFolder = Path.Combine(Path.GetTempPath(), "NetworkShare");
        Directory.CreateDirectory(networkFolder);

        // Paths for the source and output documents.
        string sourcePath = Path.Combine(networkFolder, "sample.docx");
        string outputPath = Path.Combine(networkFolder, "sample_watermarked.docx");

        // Create a blank document and save it to the simulated network location.
        Document blankDoc = new Document();
        blankDoc.Save(sourcePath);

        // Open the document from the network share using a FileStream to ensure handles are disposed.
        using (FileStream stream = new FileStream(sourcePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
        {
            Document doc = new Document(stream);

            // Add a text watermark.
            doc.Watermark.SetText("Confidential");

            // Save the watermarked document.
            doc.Save(outputPath);
        }

        // Verify that the output file was created (no console output required).
        bool created = File.Exists(outputPath);
    }
}
