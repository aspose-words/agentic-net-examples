using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing several headings.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;

        builder.Writeln("Chapter 1");
        builder.Writeln("Chapter 2");
        builder.Writeln("Chapter 3");

        // Save the original document (optional, just for reference).
        doc.Save("input.docx");

        // Replace each occurrence of the word "Chapter" with itself followed by a page break.
        // The form‑feed character (\f) is interpreted by Aspose.Words as a page break.
        int replacedCount = doc.Range.Replace("Chapter", "Chapter\f", new FindReplaceOptions());

        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one heading to be replaced.");

        // Save the modified document.
        doc.Save("output.docx");

        // Verify that the output file was created.
        if (!File.Exists("output.docx"))
            throw new FileNotFoundException("The output document was not created.");
    }
}
