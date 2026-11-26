using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Simulate a network share by creating a temporary folder.
        string networkSharePath = Path.Combine(Path.GetTempPath(), "NetworkShare");
        Directory.CreateDirectory(networkSharePath);

        // Path to the source document located on the simulated network share.
        string sourceDocPath = Path.Combine(networkSharePath, "source.docx");

        // Create a simple source document if it does not already exist.
        if (!File.Exists(sourceDocPath))
        {
            var doc = new Document();
            var builder = new DocumentBuilder(doc);
            builder.Writeln("This is a sample document.");
            doc.Save(sourceDocPath);
        }

        // Path for the output document with the watermark.
        string outputDocPath = Path.Combine(networkSharePath, "watermarked.docx");

        // Open the document from the network share using a FileStream to ensure proper disposal.
        using (FileStream stream = new FileStream(sourceDocPath, FileMode.Open, FileAccess.ReadWrite, FileShare.Read))
        {
            var doc = new Document(stream);

            // Add a text watermark to the document.
            doc.Watermark.SetText("CONFIDENTIAL");

            // Save the watermarked document.
            doc.Save(outputDocPath);
        }

        // The program ends here; no user interaction is required.
    }
}
