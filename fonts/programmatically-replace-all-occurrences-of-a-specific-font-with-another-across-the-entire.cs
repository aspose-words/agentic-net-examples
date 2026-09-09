using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some paragraphs using the font that we want to replace.
        builder.Font.Name = "Courier New";               // Old font name.
        builder.Writeln("First paragraph using the old font.");
        builder.Writeln("Second paragraph also using the old font.");

        // Add a paragraph that uses a different font – it should stay unchanged.
        builder.Font.Name = "Times New Roman";
        builder.Writeln("Paragraph with a different font that must remain.");

        // Prepare find‑replace options that will apply a new font to matched text.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ApplyFont.Name = "Arial"; // New font name to replace the old one.

        // Use a regular expression that matches every character in the document.
        // The replacement string "$0" keeps the original text unchanged,
        // while the options apply the new font to the matched runs.
        Regex allText = new Regex("(?s).+"); // (?s) enables single‑line mode.
        doc.Range.Replace(allText, "$0", options);

        // Save the resulting document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "FontReplaced.docx");
        doc.Save(outputPath);
    }
}
