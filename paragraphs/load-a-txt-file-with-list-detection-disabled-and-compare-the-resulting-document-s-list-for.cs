using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Prepare sample TXT content with list‑like lines.
        string txtContent = "1. First item\r\n2. Second item\r\nJust a normal paragraph.";
        string tempFolder = Path.Combine(Path.GetTempPath(), "AsposeWordsExample");
        Directory.CreateDirectory(tempFolder);
        string txtPath = Path.Combine(tempFolder, "sample.txt");
        File.WriteAllText(txtPath, txtContent);

        // Load the TXT file with default list detection (enabled).
        var docEnabled = new Document(txtPath);
        docEnabled.Save(Path.Combine(tempFolder, "doc_enabled.docx"));

        // Create a copy that simulates list detection disabled by removing any detected list formatting.
        var docDisabled = (Document)docEnabled.Clone(true);
        var paragraphs = docDisabled.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in paragraphs)
        {
            if (para.ListFormat.IsListItem)
                para.ListFormat.RemoveNumbers(); // Remove list formatting to mimic detection disabled.
        }
        docDisabled.Save(Path.Combine(tempFolder, "doc_disabled.docx"));

        // Helper to get list‑item flags for a document.
        static bool[] GetListFlags(Document doc)
        {
            var paras = doc.GetChildNodes(NodeType.Paragraph, true);
            bool[] flags = new bool[paras.Count];
            for (int i = 0; i < paras.Count; i++)
            {
                var para = (Paragraph)paras[i];
                flags[i] = para.ListFormat.IsListItem;
            }
            return flags;
        }

        bool[] disabledFlags = GetListFlags(docDisabled);
        bool[] enabledFlags = GetListFlags(docEnabled);
        int paragraphCount = Math.Max(disabledFlags.Length, enabledFlags.Length);

        var enabledParagraphs = docEnabled.GetChildNodes(NodeType.Paragraph, true);

        // Build comparison report.
        using var report = new StringWriter();

        report.WriteLine("Paragraph list detection comparison:");
        report.WriteLine();

        for (int i = 0; i < paragraphCount; i++)
        {
            string paraText = i < enabledParagraphs.Count
                ? ((Paragraph)enabledParagraphs[i]).GetText().TrimEnd('\r', '\n')
                : "(no paragraph)";

            bool isListDisabled = i < disabledFlags.Length && disabledFlags[i];
            bool isListEnabled = i < enabledFlags.Length && enabledFlags[i];

            report.WriteLine($"Paragraph {i + 1}: \"{paraText}\"");
            report.WriteLine($"  List detection disabled: {(isListDisabled ? "List item" : "Not a list")}");
            report.WriteLine($"  List detection enabled : {(isListEnabled ? "List item" : "Not a list")}");
            report.WriteLine();
        }

        // Save the report.
        string reportPath = Path.Combine(tempFolder, "list_comparison.txt");
        File.WriteAllText(reportPath, report.ToString());

        // Output paths for reference.
        Console.WriteLine("Comparison report saved to: " + reportPath);
        Console.WriteLine("Documents saved to: " + tempFolder);
    }
}
