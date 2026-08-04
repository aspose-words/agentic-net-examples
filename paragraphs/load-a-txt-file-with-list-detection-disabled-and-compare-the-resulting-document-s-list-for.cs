using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Prepare sample plain‑text content that contains list items.
        // The first three lines use a dot delimiter (e.g., "1.") – always detected as a list.
        // The next three lines use a whitespace delimiter (e.g., "1 ") – detected only when
        // DetectNumberingWithWhitespaces is true.
        string txtContent =
            "1. First item\r\n" +
            "2. Second item\r\n" +
            "3. Third item\r\n\r\n" +
            "1 Fourth item\r\n" +
            "2 Fourth item\r\n" +
            "3 Fourth item";

        // Write the content to a temporary file.
        string txtPath = Path.Combine(Directory.GetCurrentDirectory(), "sample.txt");
        File.WriteAllText(txtPath, txtContent);

        // -----------------------------------------------------------------
        // Load with list detection enabled (default settings).
        // -----------------------------------------------------------------
        Document docEnabled = new Document(txtPath, new TxtLoadOptions());

        // Count paragraphs that are recognized as list items.
        int enabledListItemCount = docEnabled
            .GetChildNodes(NodeType.Paragraph, true)
            .Cast<Paragraph>()
            .Count(p => p.IsListItem);

        // Save the document for visual verification (optional).
        string enabledDocPath = Path.Combine(Directory.GetCurrentDirectory(), "enabled.docx");
        docEnabled.Save(enabledDocPath);

        // -----------------------------------------------------------------
        // Load with list detection disabled.
        // -----------------------------------------------------------------
        TxtLoadOptions disabledOptions = new TxtLoadOptions
        {
            // Turn off automatic numbering detection completely.
            AutoNumberingDetection = false,
            // Also disable whitespace‑delimited list detection.
            DetectNumberingWithWhitespaces = false
        };
        Document docDisabled = new Document(txtPath, disabledOptions);

        int disabledListItemCount = docDisabled
            .GetChildNodes(NodeType.Paragraph, true)
            .Cast<Paragraph>()
            .Count(p => p.IsListItem);

        // Save the document for visual verification (optional).
        string disabledDocPath = Path.Combine(Directory.GetCurrentDirectory(), "disabled.docx");
        docDisabled.Save(disabledDocPath);

        // -----------------------------------------------------------------
        // Output the comparison results.
        // -----------------------------------------------------------------
        Console.WriteLine($"List items with detection enabled : {enabledListItemCount}");
        Console.WriteLine($"List items with detection disabled: {disabledListItemCount}");
    }
}
