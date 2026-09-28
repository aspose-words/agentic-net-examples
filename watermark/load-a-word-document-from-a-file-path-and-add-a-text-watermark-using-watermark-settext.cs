using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample source document.
        string inputPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document.");
        doc.Save(inputPath);

        // Load the document from the file path.
        Document loadedDoc = new Document(inputPath);

        // Add a text watermark.
        loadedDoc.Watermark.SetText("CONFIDENTIAL");

        // Save the document with the watermark.
        string outputPath = "output.docx";
        loadedDoc.Save(outputPath);

        // Validate that the output file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Watermark added and document saved to " + outputPath);
        }
    }
}
