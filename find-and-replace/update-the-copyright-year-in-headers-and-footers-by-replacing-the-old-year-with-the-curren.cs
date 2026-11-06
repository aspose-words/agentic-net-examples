using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a sample document with header and footer containing an old year.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Header with old year.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("© 2022 My Company");

        // Footer with old year.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("© 2022 My Company - All rights reserved");

        // Body content.
        builder.MoveToDocumentEnd();
        builder.Writeln("Document body content.");

        // Save the initial document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document as if it were an existing file.
        Document loaded = new Document(inputPath);

        // Define the old year to replace and the new current year.
        string oldYear = "2022";
        string newYear = DateTime.Now.Year.ToString();

        // Perform the replacement throughout the document (including headers/footers).
        int replacedCount = loaded.Range.Replace(oldYear, newYear, new FindReplaceOptions());

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);

        // Create a simple JSON report of the operation.
        var report = new
        {
            Replacements = replacedCount,
            OldYear = oldYear,
            NewYear = newYear,
            OutputDocument = outputPath
        };
        string jsonReport = JsonConvert.SerializeObject(report, Formatting.Indented);
        File.WriteAllText("report.json", jsonReport);
    }
}
