using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Create a new document and add sample text containing dates.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Important dates:");
        builder.Writeln("Project start: January 1, 2020");
        builder.Writeln("Release: March 15, 2021");
        builder.Writeln("End of support: December 31, 2022");

        // Configure find‑replace options to use a callback for custom replacement logic.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new MonthNameReplacer(),
            // Use the .NET regex engine (not the legacy Aspose engine).
            LegacyMode = false
        };

        // Regular expression that matches "MonthName day, year".
        Regex datePattern = new Regex(
            @"\b(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{1,2}),\s+(\d{4})\b",
            RegexOptions.IgnoreCase);

        // Perform the replacement across the whole document range.
        doc.Range.Replace(datePattern, string.Empty, options);

        // Save the modified document.
        doc.Save("ConvertedDates.docx");
    }
}

// Callback that converts a matched month name to its numeric representation
// and formats the date as "MM/dd/yyyy".
class MonthNameReplacer : IReplacingCallback
{
    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // Groups: 1 = month name, 2 = day, 3 = year.
        string monthName = args.Match.Groups[1].Value;
        string day = args.Match.Groups[2].Value;
        string year = args.Match.Groups[3].Value;

        int monthNumber = MonthNameToNumber(monthName);
        // Build the replacement string with leading zeros for month and day.
        string replacement = $"{monthNumber:D2}/{int.Parse(day):D2}/{year}";
        args.Replacement = replacement;

        return ReplaceAction.Replace;
    }

    private int MonthNameToNumber(string month)
    {
        switch (month.ToLower())
        {
            case "january":   return 1;
            case "february":  return 2;
            case "march":     return 3;
            case "april":     return 4;
            case "may":       return 5;
            case "june":      return 6;
            case "july":      return 7;
            case "august":    return 8;
            case "september": return 9;
            case "october":   return 10;
            case "november":  return 11;
            case "december":  return 12;
            default:          return 0;
        }
    }
}
