using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with date strings in the format "January 1, 2020".
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The conference starts on January 1, 2020 and ends on February 12, 2021.");
        builder.Writeln("Another date: March 3, 2022.");

        // Regular expression to match month name, day and year.
        Regex dateRegex = new Regex(@"\b(January|February|March|April|May|June|July|August|September|October|November|December) (\d{1,2}), (\d{4})\b");

        // Set up find‑replace options with a custom callback to format the date.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new DateReplacer();

        // Perform the replacement. The replacement string argument is ignored when using a callback.
        int replacedCount = doc.Range.Replace(dateRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one date replacement, but none were made.");

        // Save the modified document.
        doc.Save("output.docx");
    }
}

// Callback that converts a matched date string to "MM/dd/yyyy" format.
public class DateReplacer : IReplacingCallback
{
    private static readonly Dictionary<string, int> MonthNumbers = new()
    {
        { "January", 1 }, { "February", 2 }, { "March", 3 }, { "April", 4 },
        { "May", 5 }, { "June", 6 }, { "July", 7 }, { "August", 8 },
        { "September", 9 }, { "October", 10 }, { "November", 11 }, { "December", 12 }
    };

    public ReplaceAction Replacing(ReplacingArgs args)
    {
        Match match = args.Match;
        string monthName = match.Groups[1].Value;
        int day = int.Parse(match.Groups[2].Value);
        int year = int.Parse(match.Groups[3].Value);

        int monthNumber = MonthNumbers[monthName];
        string formattedDate = $"{monthNumber:D2}/{day:D2}/{year}";

        args.Replacement = formattedDate;
        return ReplaceAction.Replace;
    }
}
