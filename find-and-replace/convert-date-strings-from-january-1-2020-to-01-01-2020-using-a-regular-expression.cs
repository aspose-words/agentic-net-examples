using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class DateReplacer : IReplacingCallback
{
    // Mapping from month name to its two‑digit number.
    private static readonly Dictionary<string, string> MonthMap = new()
    {
        { "January",   "01" },
        { "February",  "02" },
        { "March",     "03" },
        { "April",     "04" },
        { "May",       "05" },
        { "June",      "06" },
        { "July",      "07" },
        { "August",    "08" },
        { "September", "09" },
        { "October",   "10" },
        { "November",  "11" },
        { "December",  "12" }
    };

    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // Extract the captured groups from the regex match.
        Match match = args.Match;
        string monthName = match.Groups[1].Value;
        int day = int.Parse(match.Groups[2].Value);
        string year = match.Groups[3].Value;

        // Build the replacement string in MM/dd/yyyy format.
        string monthNumber = MonthMap[monthName];
        string dayString = day.ToString("D2");
        string replacementText = $"{monthNumber}/{dayString}/{year}";

        // Assign the replacement text. Using a string keeps the code compatible with all Aspose.Words versions.
        args.Replacement = replacementText;
        return ReplaceAction.Replace;
    }
}

public class Program
{
    public static void Main()
    {
        // Create a sample document containing dates in the long format.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The first date is January 1, 2020.");
        builder.Writeln("Another date: February 12, 2021.");
        builder.Writeln("And March 5, 2022 appears here.");

        // Regular expression that captures month name, day and year.
        Regex dateRegex = new Regex(@"\b(January|February|March|April|May|June|July|August|September|October|November|December) (\d{1,2}), (\d{4})\b");

        // Set up find‑replace options with the custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new DateReplacer()
        };

        // Perform the replacement using the regex and callback.
        int replacedCount = doc.Range.Replace(dateRegex, string.Empty, options);
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one date replacement.");

        // Save the modified document.
        doc.Save("output.docx");
    }
}
