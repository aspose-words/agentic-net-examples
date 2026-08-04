using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a blank Word document.
        Document doc = new Document();

        // Apply a text watermark to the document.
        doc.Watermark.SetText("Confidential");

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Watermarked.docx");

        // Save the document as DOCX.
        doc.Save(outputPath);

        // Validate that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Watermark applied and document saved successfully.");
        }
    }
}
