using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Paths for the input and output documents.
        string inputPath = "input.docx";
        string outputPath = "output.docx";

        // Create a sample source document if it does not already exist.
        if (!File.Exists(inputPath))
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln("This is a sample document.");
            sampleDoc.Save(inputPath);
        }

        // Load the existing document.
        Document doc = new Document(inputPath);

        // Add a text watermark.
        doc.Watermark.SetText("CONFIDENTIAL");

        // Save the document with the watermark applied.
        doc.Save(outputPath);
    }
}
