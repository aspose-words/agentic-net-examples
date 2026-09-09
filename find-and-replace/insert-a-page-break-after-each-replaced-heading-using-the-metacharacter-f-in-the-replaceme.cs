using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample headings that we will replace.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter One");
        builder.Writeln("Chapter Two");
        builder.Writeln("Chapter Three");

        // Replace each heading text with the same text followed by a page break.
        // The form‑feed character (\f) is interpreted by Aspose.Words as a page break.
        FindReplaceOptions options = new FindReplaceOptions();
        int replacedCount = doc.Range.Replace("Chapter", "Chapter\f", options);

        // Ensure that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No headings were replaced.");

        // Save the modified document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");
        doc.Save(outputPath);
    }
}
