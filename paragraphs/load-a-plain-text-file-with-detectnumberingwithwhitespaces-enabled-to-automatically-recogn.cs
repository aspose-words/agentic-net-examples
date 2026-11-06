using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Prepare a temporary plain‑text file containing list items with leading whitespaces.
        string tempFolder = Path.Combine(Path.GetTempPath(), "AsposeWordsDemo");
        Directory.CreateDirectory(tempFolder);
        string txtPath = Path.Combine(tempFolder, "sample.txt");

        string[] lines =
        {
            " 1. First item",
            " 2. Second item",
            "   a) Sub item",
            "   b) Another sub item",
            "Plain paragraph without list."
        };
        File.WriteAllLines(txtPath, lines);

        // Load the text file with DetectNumberingWithWhitespaces enabled.
        // Use TxtLoadOptions because DetectNumberingWithWhitespaces is defined there.
        TxtLoadOptions loadOptions = new TxtLoadOptions
        {
            DetectNumberingWithWhitespaces = true
        };
        Document doc = new Document(txtPath, loadOptions);

        // Count paragraphs that were recognized as list items.
        int listItemCount = 0;
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.ListFormat.IsListItem)
                listItemCount++;
        }

        // Save the resulting document.
        string outPath = Path.Combine(tempFolder, "output.docx");
        doc.Save(outPath);

        // Output the result (no interactive prompts).
        Console.WriteLine($"List items detected: {listItemCount}");
        Console.WriteLine($"Document saved to: {outPath}");
    }
}
