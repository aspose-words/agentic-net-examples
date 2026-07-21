using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Prepare sample plain‑text content containing numbered list items.
        string txtContent =
            "1. First item\n" +
            "2. Second item\n" +
            "3. Third item\n\n" +
            "1 Fourth item\n" +
            "2 Fourth item\n" +
            "3 Fourth item";

        // Define file paths in the current working directory.
        string txtPath = Path.Combine(Directory.GetCurrentDirectory(), "Sample.txt");
        string enabledDocPath = Path.Combine(Directory.GetCurrentDirectory(), "Enabled.docx");
        string disabledDocPath = Path.Combine(Directory.GetCurrentDirectory(), "Disabled.docx");

        // Write the sample text to a file.
        File.WriteAllText(txtPath, txtContent);

        // ---------- Load with list detection enabled (default options) ----------
        TxtLoadOptions loadOptionsEnabled = new TxtLoadOptions(); // DetectNumberingWithWhitespaces = true, AutoNumberingDetection = true
        Document docEnabled = new Document(txtPath, loadOptionsEnabled);

        // Count paragraphs that are recognized as list items.
        int enabledListCount = docEnabled
            .GetChildNodes(NodeType.Paragraph, true)
            .Cast<Paragraph>()
            .Count(p => p.IsListItem);

        // Save the document for visual inspection (optional).
        docEnabled.Save(enabledDocPath);

        // ---------- Load with list detection disabled ----------
        TxtLoadOptions loadOptionsDisabled = new TxtLoadOptions
        {
            AutoNumberingDetection = false,          // Turn off automatic numbering detection.
            DetectNumberingWithWhitespaces = false   // Do not treat whitespace as a list delimiter.
        };
        Document docDisabled = new Document(txtPath, loadOptionsDisabled);

        int disabledListCount = docDisabled
            .GetChildNodes(NodeType.Paragraph, true)
            .Cast<Paragraph>()
            .Count(p => p.IsListItem);

        // Save the document loaded with detection disabled.
        docDisabled.Save(disabledDocPath);

        // Output the comparison results.
        Console.WriteLine($"List items detected with detection enabled : {enabledListCount}");
        Console.WriteLine($"List items detected with detection disabled: {disabledListCount}");
    }
}
