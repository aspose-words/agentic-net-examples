using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define file paths.
        string inputPath = "Sample.docx";
        string outputPath = "Watermarked.docx";

        // Create a simple blank document if it does not already exist.
        if (!File.Exists(inputPath))
        {
            Document blankDoc = new Document();
            blankDoc.Save(inputPath);
        }

        // Load the existing document.
        Document doc = new Document(inputPath);

        // Add a text watermark.
        doc.Watermark.SetText("CONFIDENTIAL");

        // Save the document with the watermark.
        doc.Save(outputPath);

        // Verify that the output file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Watermark applied successfully. Output saved to '{outputPath}'.");
        }
    }
}
