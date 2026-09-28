using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Simulate a network share by using a temporary folder.
        string networkSharePath = Path.Combine(Path.GetTempPath(), "NetworkShare");
        Directory.CreateDirectory(networkSharePath);

        // Path for the source document on the "network share".
        string sourceDocPath = Path.Combine(networkSharePath, "source.docx");
        // Path for the output document with the watermark.
        string outputDocPath = Path.Combine(networkSharePath, "output.docx");

        // Create a simple sample document and save it to the network share location.
        var sampleDoc = new Document();
        var builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("Hello World!");
        sampleDoc.Save(sourceDocPath);

        // Open the document from the network share using a FileStream to ensure proper handle disposal.
        using (FileStream stream = new FileStream(sourceDocPath, FileMode.Open, FileAccess.ReadWrite, FileShare.Read))
        {
            // Load the document from the stream.
            var doc = new Document(stream);

            // Add a text watermark.
            doc.Watermark.SetText("CONFIDENTIAL");

            // Save the watermarked document back to the network share.
            doc.Save(outputDocPath);
        }

        // Simple validation that the output file was created.
        if (File.Exists(outputDocPath))
        {
            Console.WriteLine("Watermark applied successfully. Output saved to: " + outputDocPath);
        }
        else
        {
            Console.WriteLine("Failed to create the output document.");
        }
    }
}
