using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new document and add text that contains several manual line breaks.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Manual line breaks are inserted with ControlChar.LineBreak.
        // This creates a single paragraph with multiple line‑break characters inside it.
        builder.Write("First line" + ControlChar.LineBreak + ControlChar.LineBreak + ControlChar.LineBreak +
                      "Second line" + ControlChar.LineBreak + ControlChar.LineBreak + "Third line");

        // Save the sample source document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for find‑and‑replace.
        Document loaded = new Document(inputPath);

        // Regex that matches two or more consecutive line‑break characters.
        Regex multipleLineBreaks = new Regex(@"(\v){2,}");

        // Replace the matched line breaks with a single paragraph break.
        // Use the Aspose.Words meta‑character "&p" to insert a paragraph break.
        int replacedCount = loaded.Range.Replace(multipleLineBreaks, "&p", new FindReplaceOptions());

        // Verify that at least one replacement was performed.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
