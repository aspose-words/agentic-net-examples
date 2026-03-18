using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Notes;

class FootnoteExtractor
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Directory where individual footnote files will be saved.
        string outputDir = "Footnotes";
        Directory.CreateDirectory(outputDir);

        // Iterate through all footnotes in the document.
        NodeCollection footnotes = doc.GetChildNodes(NodeType.Footnote, true);
        for (int i = 0; i < footnotes.Count; i++)
        {
            Footnote footnote = (Footnote)footnotes[i];

            // Get the plain text of the footnote (includes all child paragraphs/tables).
            string footnoteText = footnote.GetText().Trim();

            // Build a file name using the footnote index (1‑based) and its reference mark.
            string fileName = $"Footnote_{i + 1}_{SanitizeFileName(footnote.ReferenceMark)}.txt";
            string filePath = Path.Combine(outputDir, fileName);

            // Write the footnote text to a separate file.
            File.WriteAllText(filePath, footnoteText);
        }
    }

    // Helper to remove invalid characters from file names.
    private static string SanitizeFileName(string name)
    {
        foreach (char c in Path.GetInvalidFileNameChars())
            name = name.Replace(c.ToString(), string.Empty);
        return string.IsNullOrWhiteSpace(name) ? "NoRef" : name;
    }
}
