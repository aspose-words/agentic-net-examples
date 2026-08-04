using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "input.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");

        // -----------------------------------------------------------------
        // 1. Create a sample document that contains multiple manual line breaks.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write text that includes several consecutive manual line‑break characters.
        // Use ControlChar.LineBreak to insert a manual line break.
        builder.Write("First line." + ControlChar.LineBreak + ControlChar.LineBreak + ControlChar.LineBreak +
                      "Second line." + ControlChar.LineBreak + ControlChar.LineBreak +
                      "Third line.");

        // Save the source document.
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // 2. Load the document and replace sequences of manual line‑breaks.
        // -----------------------------------------------------------------
        Document loaded = new Document(inputPath);

        // Regex that matches two or more consecutive manual line‑break characters.
        // ControlChar.LineBreak is a string representing the line‑break character.
        string lineBreak = ControlChar.LineBreak;
        Regex multipleLineBreaks = new Regex($"{Regex.Escape(lineBreak)}{{2,}}");

        // Replace each match with a single paragraph break.
        // Use the meta‑character "&p" which Aspose.Words interprets as a paragraph break.
        int replacedCount = loaded.Range.Replace(multipleLineBreaks, "&p", new FindReplaceOptions());

        // Ensure that at least one replacement was performed.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one line‑break replacement.");

        // -----------------------------------------------------------------
        // 3. Save the modified document.
        // -----------------------------------------------------------------
        loaded.Save(outputPath);
    }
}
