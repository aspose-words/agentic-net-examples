using System;
using System.IO;
using System.Text;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Sample text containing different list delimiters.
        string text = "Full stop delimiters:\n" +
                      "1. First list item 1\n" +
                      "2. First list item 2\n" +
                      "3. First list item 3\n\n" +
                      "Right bracket delimiters:\n" +
                      "1) Second list item 1\n" +
                      "2) Second list item 2\n" +
                      "3) Second list item 3\n\n" +
                      "Bullet delimiters:\n" +
                      "• Third list item 1\n" +
                      "• Third list item 2\n" +
                      "• Third list item 3\n\n" +
                      "Whitespace delimiters:\n" +
                      "1 Fourth list item 1\n" +
                      "2 Fourth list item 2\n" +
                      "3 Fourth list item 3";

        // Create a temporary folder for the example files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Write the sample text to a temporary TXT file.
        string txtPath = Path.Combine(outputDir, "sample.txt");
        File.WriteAllText(txtPath, text, Encoding.UTF8);

        // Load the TXT file with list detection enabled (default behavior).
        var loadOptionsEnabled = new TxtLoadOptions
        {
            DetectNumberingWithWhitespaces = true
        };
        Document docEnabled = new Document(txtPath, loadOptionsEnabled);

        // Load the same TXT file with list detection disabled.
        var loadOptionsDisabled = new TxtLoadOptions
        {
            DetectNumberingWithWhitespaces = false
        };
        Document docDisabled = new Document(txtPath, loadOptionsDisabled);

        // Analyze the number of lists detected in each document.
        int listsEnabled = docEnabled.Lists.Count;
        int listsDisabled = docDisabled.Lists.Count;

        // Determine whether the "Fourth list" paragraph is recognized as a list item.
        bool fourthListIsListEnabled = docEnabled.GetChildNodes(NodeType.Paragraph, true)
            .OfType<Paragraph>()
            .Any(p => p.GetText().Contains("Fourth list") && p.ListFormat.IsListItem);

        bool fourthListIsListDisabled = docDisabled.GetChildNodes(NodeType.Paragraph, true)
            .OfType<Paragraph>()
            .Any(p => p.GetText().Contains("Fourth list") && p.ListFormat.IsListItem);

        // Save the resulting documents for manual inspection if needed.
        docEnabled.Save(Path.Combine(outputDir, "enabled.docx"));
        docDisabled.Save(Path.Combine(outputDir, "disabled.docx"));

        // Output the comparison results.
        Console.WriteLine($"Lists count with detection enabled: {listsEnabled}");
        Console.WriteLine($"Lists count with detection disabled: {listsDisabled}");
        Console.WriteLine($"\"Fourth list\" recognized as list with detection enabled: {fourthListIsListEnabled}");
        Console.WriteLine($"\"Fourth list\" recognized as list with detection disabled: {fourthListIsListDisabled}");
    }
}
