using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a sample document with multiple paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First paragraph.");
        builder.Writeln("Second paragraph to copy.");
        builder.Writeln("Third paragraph.");

        const string inputPath = "sample.docx";
        doc.Save(inputPath);

        // Load the document from the file.
        Document loaded = new Document(inputPath);

        // Extract the second paragraph (index 1).
        Paragraph paragraph = loaded.FirstSection?.Body?.Paragraphs[1];
        if (paragraph == null)
            throw new InvalidOperationException("Target paragraph not found.");

        string extractedText = paragraph.GetText().TrimEnd('\r', '\n');

        // Write the extracted text to a file for verification (simulating clipboard copy).
        const string outputPath = "extracted.txt";
        File.WriteAllText(outputPath, extractedText);

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Extracted output file was not created.");

        // Indicate successful completion.
        Console.WriteLine("Extraction completed and text written to '" + outputPath + "'.");
    }
}
