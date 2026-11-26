using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a sample document with plain‑text numbered list items.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("This is a normal paragraph.");
        builder.Writeln("1. First item");
        builder.Writeln("2. Second item");
        builder.Writeln("3. Third item");
        builder.Writeln("Another normal paragraph.");

        // Save the initial document.
        const string sourcePath = "PlainList.docx";
        doc.Save(sourcePath);

        // Load the document for processing.
        Document loadedDoc = new Document(sourcePath);

        // Create a numbered list definition.
        List numberedList = loadedDoc.Lists.Add(ListTemplate.NumberDefault);

        // Regular expression to detect plain‑text numbering (e.g., "1. Item").
        Regex listRegex = new Regex(@"^\s*(\d+)\.\s+(.*)", RegexOptions.Compiled);

        // Iterate over all paragraphs in the document.
        foreach (Paragraph paragraph in loadedDoc.GetChildNodes(NodeType.Paragraph, true))
        {
            string paragraphText = paragraph.GetText().TrimEnd('\r', '\a');

            Match match = listRegex.Match(paragraphText);
            if (match.Success)
            {
                // Extract the text after the numbering.
                string cleanText = match.Groups[2].Value;

                // Remove existing runs and add a new run with the cleaned text.
                paragraph.Runs.Clear();
                Run run = new Run(loadedDoc, cleanText);
                paragraph.AppendChild(run);

                // Apply the numbered list format to the paragraph.
                paragraph.ListFormat.List = numberedList;
                paragraph.ListFormat.ListLevelNumber = 0;
            }
        }

        // Save the transformed document.
        const string outputPath = "ConvertedList.docx";
        loadedDoc.Save(outputPath);
    }
}
