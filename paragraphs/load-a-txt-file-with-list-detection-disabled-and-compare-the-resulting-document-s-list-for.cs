using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Prepare a sample plain‑text file containing numbered and bulleted list items.
        string txtContent =
            "1. First numbered item\r\n" +
            "2. Second numbered item\r\n" +
            "3. Third numbered item\r\n" +
            "A regular paragraph without list formatting.\r\n" +
            "- First bullet item\r\n" +
            "- Second bullet item\r\n";

        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        string txtPath = Path.Combine(artifactsDir, "Sample.txt");
        File.WriteAllText(txtPath, txtContent);

        // Load the text file with the default options (list detection enabled).
        Document docWithDetection = new Document(txtPath);

        // Load the same text file with list detection disabled.
        TxtLoadOptions loadOptions = new TxtLoadOptions
        {
            // Disables automatic numbering detection while loading plain‑text.
            AutoNumberingDetection = false
        };
        Document docWithoutDetection = new Document(txtPath, loadOptions);

        // Count paragraphs that are recognized as list items in each document.
        int countWithDetection = docWithDetection
            .GetChildNodes(NodeType.Paragraph, true)
            .Cast<Paragraph>()
            .Count(p => p.IsListItem);

        int countWithoutDetection = docWithoutDetection
            .GetChildNodes(NodeType.Paragraph, true)
            .Cast<Paragraph>()
            .Count(p => p.IsListItem);

        // Output the comparison result.
        Console.WriteLine($"List items with detection enabled : {countWithDetection}");
        Console.WriteLine($"List items with detection disabled: {countWithoutDetection}");

        // Save both documents so the difference can be inspected manually if needed.
        docWithDetection.Save(Path.Combine(artifactsDir, "WithListDetection.docx"));
        docWithoutDetection.Save(Path.Combine(artifactsDir, "WithoutListDetection.docx"));
    }
}
