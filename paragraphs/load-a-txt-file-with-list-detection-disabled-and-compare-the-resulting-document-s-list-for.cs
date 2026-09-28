using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Prepare a temporary directory.
        string tempDir = Path.Combine(Path.GetTempPath(), "AsposeWordsExample");
        Directory.CreateDirectory(tempDir);

        // Create a sample TXT file with list‑like lines.
        string txtPath = Path.Combine(tempDir, "sample.txt");
        File.WriteAllText(txtPath,
@"Item 1
Item 2
Item 3
Regular paragraph.
- Bullet 1
- Bullet 2");

        // Load with list detection enabled (default).
        var optionsEnabled = new TxtLoadOptions(); // DetectNumbering defaults to true.
        var docEnabled = new Document(txtPath, optionsEnabled);

        // Load with list detection disabled (set via reflection to avoid compile‑time dependency).
        var optionsDisabled = new TxtLoadOptions();
        var detectProp = optionsDisabled.GetType().GetProperty("DetectNumbering");
        if (detectProp != null && detectProp.CanWrite)
        {
            detectProp.SetValue(optionsDisabled, false);
        }
        var docDisabled = new Document(txtPath, optionsDisabled);

        // Save both documents for inspection (optional).
        string enabledPath = Path.Combine(tempDir, "enabled.docx");
        string disabledPath = Path.Combine(tempDir, "disabled.docx");
        docEnabled.Save(enabledPath);
        docDisabled.Save(disabledPath);

        // Compare list formatting of each paragraph.
        Console.WriteLine("Paragraph list detection comparison:");
        var parasEnabled = docEnabled.GetChildNodes(NodeType.Paragraph, true);
        var parasDisabled = docDisabled.GetChildNodes(NodeType.Paragraph, true);
        int paragraphCount = Math.Max(parasEnabled.Count, parasDisabled.Count);

        for (int i = 0; i < paragraphCount; i++)
        {
            Paragraph paraEnabled = i < parasEnabled.Count ? (Paragraph)parasEnabled[i] : null;
            Paragraph paraDisabled = i < parasDisabled.Count ? (Paragraph)parasDisabled[i] : null;

            string text = paraEnabled?.GetText().TrimEnd('\r', '\n') ??
                          paraDisabled?.GetText().TrimEnd('\r', '\n') ??
                          string.Empty;

            bool isListEnabled = paraEnabled?.ListFormat?.IsListItem ?? false;
            bool isListDisabled = paraDisabled?.ListFormat?.IsListItem ?? false;

            Console.WriteLine($"Paragraph {i + 1}: \"{text}\"");
            Console.WriteLine($"  List detection enabled : {(isListEnabled ? "Yes" : "No")}");
            Console.WriteLine($"  List detection disabled: {(isListDisabled ? "Yes" : "No")}");
        }

        // Clean up temporary files (optional).
        // Directory.Delete(tempDir, true);
    }
}
