using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for output files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a sample plain‑text file that contains numbered list items.
        string txtContent =
            "1. First item" + Environment.NewLine +
            "2. Second item" + Environment.NewLine +
            "3. Third item" + Environment.NewLine +
            Environment.NewLine +
            "A normal paragraph without list formatting.";
        string txtPath = Path.Combine(artifactsDir, "Sample.txt");
        File.WriteAllText(txtPath, txtContent);

        // Load the text file with default list detection (enabled).
        Document docWithLists = new Document(txtPath);

        // Load the same text file with automatic list detection turned off.
        TxtLoadOptions loadOptions = new TxtLoadOptions { AutoNumberingDetection = false };
        Document docWithoutLists = new Document(txtPath, loadOptions);

        // Count paragraphs that are recognized as list items in each document.
        int enabledListCount = docWithLists
            .GetChildNodes(NodeType.Paragraph, true)
            .Cast<Paragraph>()
            .Count(p => p.IsListItem);

        int disabledListCount = docWithoutLists
            .GetChildNodes(NodeType.Paragraph, true)
            .Cast<Paragraph>()
            .Count(p => p.IsListItem);

        // Save both documents for visual inspection.
        docWithLists.Save(Path.Combine(artifactsDir, "Enabled.docx"));
        docWithoutLists.Save(Path.Combine(artifactsDir, "Disabled.docx"));

        // Output the comparison result.
        Console.WriteLine($"List items detected with default settings: {enabledListCount}");
        Console.WriteLine($"List items detected with AutoNumberingDetection = false: {disabledListCount}");
    }
}
