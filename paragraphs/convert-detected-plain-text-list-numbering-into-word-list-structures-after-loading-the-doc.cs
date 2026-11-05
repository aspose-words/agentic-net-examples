using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing plain‑text numbered items.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Introduction paragraph.");
        builder.Writeln("1. First item");
        builder.Writeln("2. Second item");
        builder.Writeln("3. Third item");
        builder.Writeln("Conclusion paragraph.");

        // Save the sample document (input).
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loadedDoc = new Document(inputPath);

        // Create a list that will be applied to detected items.
        List list = loadedDoc.Lists.Add(ListTemplate.NumberDefault);

        // Regex to detect plain‑text numbering like "1. Item".
        Regex plainNumberRegex = new Regex(@"^\s*\d+\.\s+(.*)", RegexOptions.Compiled);

        // Traverse all paragraphs and convert matching ones to list items.
        foreach (Paragraph para in loadedDoc.GetChildNodes(NodeType.Paragraph, true))
        {
            string paragraphText = para.GetText(); // Includes the paragraph break.
            string trimmed = paragraphText.TrimEnd('\r', '\n');

            Match match = plainNumberRegex.Match(trimmed);
            if (match.Success)
            {
                // Extract the actual list item text without the plain number.
                string itemText = match.Groups[1].Value;

                // Remove existing runs (the original text) and insert the clean text.
                para.Runs.Clear();
                para.Runs.Add(new Run(loadedDoc, itemText));

                // Apply list formatting.
                para.ListFormat.List = list;
                para.ListFormat.ListLevelNumber = 0;
            }
        }

        // Save the transformed document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);
    }
}
