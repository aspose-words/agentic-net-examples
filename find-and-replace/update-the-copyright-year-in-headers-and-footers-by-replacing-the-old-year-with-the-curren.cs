using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;
using System.Collections.Generic;

public class Program
{
    public static void Main()
    {
        // Define file names.
        const string inputPath = "sample.docx";
        const string outputPath = "updated.docx";

        // Step 1: Create a sample document with header and footer containing an old year.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add header with old year.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("© 2022 My Company");

        // Add footer with old year.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("Confidential - 2022");

        // Return to main body and add some content.
        builder.MoveToDocumentEnd();
        builder.Writeln("This is a sample document.");
        builder.Writeln("All rights reserved.");

        // Save the sample document.
        doc.Save(inputPath);

        // Step 2: Load the document and replace the old year with the current year.
        Document loaded = new Document(inputPath);

        string oldYear = "2022";
        string newYear = DateTime.Now.Year.ToString();

        // Perform the replacement in the whole document (including headers/footers).
        int replacedCount = loaded.Range.Replace(oldYear, newYear, new FindReplaceOptions());

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the updated document.
        loaded.Save(outputPath);

        // Optional: Write a simple JSON report of the operation.
        var report = new
        {
            InputFile = Path.GetFullPath(inputPath),
            OutputFile = Path.GetFullPath(outputPath),
            OldYear = oldYear,
            NewYear = newYear,
            ReplacementsMade = replacedCount,
            Timestamp = DateTime.UtcNow
        };

        string jsonReport = JsonConvert.SerializeObject(report, Formatting.Indented);
        File.WriteAllText("report.json", jsonReport);
    }
}
