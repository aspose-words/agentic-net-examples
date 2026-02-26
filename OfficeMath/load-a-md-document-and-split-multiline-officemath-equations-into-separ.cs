using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class SplitMultilineOfficeMath
{
    static void Main()
    {
        // Paths to the input and output Markdown files.
        string inputPath = @"C:\Docs\Input.md";
        string outputPath = @"C:\Docs\Output.md";

        // Load the Markdown document with options that preserve empty lines (optional).
        MarkdownLoadOptions loadOptions = new MarkdownLoadOptions
        {
            PreserveEmptyLines = true
        };
        Document doc = new Document(inputPath, loadOptions);

        // Save the document back to Markdown, exporting OfficeMath as MarkItDown (LaTeX compatible).
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            OfficeMathExportMode = MarkdownOfficeMathExportMode.MarkItDown
        };
        doc.Save(outputPath, saveOptions);

        // Read the saved Markdown text.
        string markdown = File.ReadAllText(outputPath, Encoding.UTF8);

        // Regex to find LaTeX blocks delimited by $$ ... $$ (including possible newlines).
        // It captures the content between the delimiters.
        string pattern = @"\$\$(?<content>[\s\S]*?)\$\$";

        // Replace each multiline block with separate single‑line blocks.
        string transformed = Regex.Replace(markdown, pattern, match =>
        {
            string content = match.Groups["content"].Value;

            // Split the content by line breaks.
            string[] lines = content.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.RemoveEmptyEntries);

            // Re‑assemble each line as its own $$ ... $$ block.
            StringBuilder sb = new StringBuilder();
            foreach (string line in lines)
            {
                if (sb.Length > 0) sb.AppendLine(); // Preserve line separation between blocks.
                sb.Append("$$");
                sb.Append(line.Trim());
                sb.Append("$$");
            }
            return sb.ToString();
        }, RegexOptions.Multiline);

        // Write the transformed Markdown back to the output file.
        File.WriteAllText(outputPath, transformed, Encoding.UTF8);
    }
}
