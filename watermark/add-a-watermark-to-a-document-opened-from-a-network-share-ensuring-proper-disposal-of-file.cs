using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for temporary files.
        string folder = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(folder);

        // Paths for the source document and the output document.
        string sourcePath = Path.Combine(folder, "source.docx");
        string outputPath = Path.Combine(folder, "watermarked.docx");

        // Create a blank document and save it to the source path.
        Document blankDoc = new Document();
        blankDoc.Save(sourcePath);

        // Open the document via a FileStream to simulate loading from a network share.
        // The using block guarantees that the file handle is released promptly.
        using (FileStream stream = new FileStream(sourcePath, FileMode.Open, FileAccess.ReadWrite, FileShare.Read))
        {
            Document doc = new Document(stream);

            // Add a text watermark to the loaded document.
            doc.Watermark.SetText("Confidential");

            // Save the watermarked document to the output path.
            doc.Save(outputPath);
        }

        // Optional verification that the output file was created.
        // No interactive prompts are used.
        if (File.Exists(outputPath))
        {
            // The file exists; processing completed successfully.
        }
    }
}
