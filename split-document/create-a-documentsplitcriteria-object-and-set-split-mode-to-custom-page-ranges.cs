using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add five pages of sample text.
        for (int i = 1; i <= 5; i++)
        {
            builder.Writeln($"Page {i}");
            if (i < 5)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Create a DocumentSplitCriteria variable and set a split mode.
        // Here we use PageBreak as an example; the variable can be used with save options if needed.
        DocumentSplitCriteria splitCriteria = DocumentSplitCriteria.PageBreak;

        // Define custom page ranges (zero‑based indices).
        // Range 1: pages 1‑2  -> start index 0, count 2
        // Range 2: pages 4‑5  -> start index 3, count 2
        var customRanges = new (int startIndex, int pageCount)[]
        {
            (0, 2),
            (3, 2)
        };

        // Prepare an output folder.
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputFolder);

        // Extract each range and save it as a separate document.
        int partNumber = 1;
        foreach (var (start, count) in customRanges)
        {
            // Extract the specified pages.
            Document partDoc = doc.ExtractPages(start, count);

            // Optionally, you could apply the split criteria to save options here.
            // For this example we simply save the extracted part.
            string outPath = Path.Combine(outputFolder, $"Part{partNumber}.docx");
            partDoc.Save(outPath, SaveFormat.Docx);

            // Verify that the file was created.
            if (!File.Exists(outPath))
                throw new InvalidOperationException($"Failed to create split document: {outPath}");

            partNumber++;
        }

        // Indicate successful completion.
        Console.WriteLine("Document split into custom page ranges completed.");
    }
}
