using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a blank document and add sample text containing dates.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The conference starts on January 1, 2020 and ends on February 12, 2021.");
        builder.Writeln("Another meeting is scheduled for March 5, 2022.");

        // Regular expression to match dates in the format "MonthName d, yyyy".
        Regex dateRegex = new Regex(@"\b(January|February|March|April|May|June|July|August|September|October|November|December) (\d{1,2}), (\d{4})\b");

        // Set up find‑replace options with a custom callback that formats the date.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new DateReplacer()
        };

        // Perform the replacement. The replacement string is ignored because the callback supplies the value.
        int replacedCount = doc.Range.Replace(dateRegex, string.Empty, options);

        if (replacedCount == 0)
            throw new InvalidOperationException("No date strings were replaced.");

        // Save the modified document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");
        doc.Save(outputPath);

        // Output the resulting text to the console (no user interaction required).
        Console.WriteLine("Replaced text:");
        Console.WriteLine(doc.GetText().Trim());
    }
}

// Callback that converts a matched date string to "MM/dd/yyyy".
public class DateReplacer : IReplacingCallback
{
    private static readonly Dictionary<string, int> MonthMap = new()
    {
        { "January", 1 }, { "February", 2 }, { "March", 3 }, { "April", 4 },
        { "May", 5 }, { "June", 6 }, { "July", 7 }, { "August", 8 },
        { "September", 9 }, { "October", 10 }, { "November", 11 }, { "December", 12 }
    };

    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // Extract month name, day and year from the match groups.
        string monthName = args.Match.Groups[1].Value;
        string dayStr = args.Match.Groups[2].Value;
        string yearStr = args.Match.Groups[3].Value;

        if (!MonthMap.TryGetValue(monthName, out int monthNumber))
            return ReplaceAction.Skip; // Unexpected month name.

        int day = int.Parse(dayStr);
        // Build the replacement string in MM/dd/yyyy format.
        args.Replacement = $"{monthNumber:D2}/{day:D2}/{yearStr}";
        return ReplaceAction.Replace;
    }
}
