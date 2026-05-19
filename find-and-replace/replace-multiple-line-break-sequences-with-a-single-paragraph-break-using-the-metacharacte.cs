using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Local file names for the sample input and output documents.
        const string inputPath = "input.docx";
        const string outputPath = "output.docx";

        // -------------------- Create a sample document --------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert text containing various line‑break characters.
        builder.Writeln("First line.");                                   // Paragraph break.
        builder.Write("Second line." + ControlChar.LineBreak);            // Manual line break (\v).
        builder.Write("Third line." + ControlChar.LineFeed);              // Line feed (\n).
        builder.Write("Fourth line." + ControlChar.CrLf);                 // CR+LF.
        builder.Writeln();                                                // Empty paragraph.
        builder.Writeln("Fifth line.");                                   // Paragraph break.

        // Save the source document.
        doc.Save(inputPath);

        // -------------------- Load the document and replace line breaks --------------------
        Document loaded = new Document(inputPath);

        // Regex that matches one or more consecutive line‑break characters,
        // including manual line break (\v) used by Aspose.Words.
        Regex lineBreakRegex = new Regex(@"(\r\n|\r|\n|\v)+");

        // Replace each matched sequence with a single paragraph break.
        // The meta‑character "&p" inserts a paragraph break.
        FindReplaceOptions options = new FindReplaceOptions();
        int replacedCount = loaded.Range.Replace(lineBreakRegex, "&p", options);

        // Validate that at least one replacement was performed.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one line break replacement.");

        // Save the modified document.
        loaded.Save(outputPath);
    }
}
