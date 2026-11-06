using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with product names.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Our product SuperWidget is great.");
        builder.Writeln("We also have superwidget in stock.");
        builder.Writeln("The SuperWidgetPro version is premium.");

        // Save the source document locally.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Configure a regular expression to match the whole word "SuperWidget"
        // ignoring case. \b ensures whole‑word matching.
        Regex regex = new Regex(@"\bSuperWidget\b", RegexOptions.IgnoreCase);

        // Perform the replacement.
        int replacedCount = loaded.Range.Replace(regex, "MegaWidget", new FindReplaceOptions());

        // Ensure that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the updated document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
