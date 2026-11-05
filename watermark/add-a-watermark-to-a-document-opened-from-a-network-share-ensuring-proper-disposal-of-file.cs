using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a temporary folder for the sample files.
        string tempFolder = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
        Directory.CreateDirectory(tempFolder);

        // Paths for the input and output documents.
        string inputPath = Path.Combine(tempFolder, "input.docx");
        string outputPath = Path.Combine(tempFolder, "output.docx");

        // Create a simple source document and save it to the input path.
        Document sourceDoc = new Document();
        sourceDoc.Save(inputPath);

        // Open the document from a "network share" using a FileStream.
        // The FileShare flags ensure the file handle is released properly.
        using (FileStream readStream = new FileStream(
            inputPath,
            FileMode.Open,
            FileAccess.ReadWrite,
            FileShare.ReadWrite))
        {
            // Load the document from the stream.
            Document doc = new Document(readStream);

            // Add a text watermark.
            doc.Watermark.SetText("CONFIDENTIAL");

            // Save the watermarked document directly to a file path.
            doc.Save(outputPath);
        }

        // Simple validation that the output file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Watermark applied and document saved to: " + outputPath);
        }
    }
}
