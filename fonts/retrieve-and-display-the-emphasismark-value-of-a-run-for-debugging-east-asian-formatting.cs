using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a run with some text.
        builder.Write("East Asian text");

        // Retrieve the run we just added.
        Run run = doc.FirstSection.Body.FirstParagraph.Runs[0];

        // Set an emphasis mark for East Asian typography.
        // Use a value that exists in the current Aspose.Words version.
        run.Font.EmphasisMark = Aspose.Words.EmphasisMark.None;

        // Output the emphasis mark value.
        Console.WriteLine($"EmphasisMark value: {run.Font.EmphasisMark}");

        // Save the document to verify output.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Confirm the file exists.
        Console.WriteLine($"Document saved successfully: {File.Exists(outputPath)}");
    }
}
