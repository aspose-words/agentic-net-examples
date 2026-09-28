using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with a page number field formatted as Roman numerals.
        builder.Writeln("Page number in Roman numerals:");
        builder.InsertField("PAGE  \\* ROMAN", "i");
        builder.Writeln();

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to {Path.GetFullPath(outputPath)}");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
