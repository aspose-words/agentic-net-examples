using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with various date range formats.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The project runs from 01-15-2021 to 02/20/2021.");
        builder.Writeln("Another period: 3/5/2020 - 4/10/2020.");
        builder.Writeln("Legacy format: 12/31/2019 - 01/15/2020.");
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Regular expression to capture date ranges in the form:
        // mm-dd-yyyy to mm/dd/yyyy, mm/dd/yyyy - mm/dd/yyyy, etc.
        Regex dateRangeRegex = new Regex(
            @"\b(?<startMonth>\d{1,2})[/-](?<startDay>\d{1,2})[/-](?<startYear>\d{4})\s*(?:to|-)\s*(?<endMonth>\d{1,2})[/-](?<endDay>\d{1,2})[/-](?<endYear>\d{4})\b",
            RegexOptions.Compiled);

        // Replacement pattern that standardizes to MM/dd/yyyy - MM/dd/yyyy.
        const string replacementPattern = "${startMonth}/${startDay}/${startYear} - ${endMonth}/${endDay}/${endYear}";

        // Configure find‑replace options to enable substitution syntax.
        FindReplaceOptions options = new FindReplaceOptions
        {
            UseSubstitutions = true
        };

        // Perform the replacement.
        int replacedCount = loaded.Range.Replace(dateRangeRegex, replacementPattern, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No date ranges were replaced. Check the regex pattern.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
