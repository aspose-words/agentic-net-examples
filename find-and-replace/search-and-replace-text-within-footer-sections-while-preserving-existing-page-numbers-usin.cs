using System;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some body content spanning two pages.
        builder.Writeln("First page content.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page content.");

        // Create a primary footer with static text and a PAGE field.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Confidential - ");
        builder.InsertField("PAGE", "?");
        builder.Write(" - Draft");

        // Replace the word "Confidential" in the footer while keeping the PAGE field intact.
        HeaderFooter footer = doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,
            FindWholeWordsOnly = false
        };
        int replaced = footer.Range.Replace("Confidential", "Public", options);
        if (replaced == 0)
            throw new InvalidOperationException("Expected at least one replacement in the footer.");

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);

        // Optional: write a simple JSON log to demonstrate the required Newtonsoft.Json package.
        var log = new { File = outputPath, ReplacementsMade = replaced };
        string jsonLog = JsonConvert.SerializeObject(log);
        Console.WriteLine(jsonLog);
    }
}
