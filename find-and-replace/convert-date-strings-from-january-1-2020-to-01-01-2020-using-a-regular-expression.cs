using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing; // Required package reference
using Newtonsoft.Json; // Required package reference

namespace FindAndReplaceDateExample
{
    // Callback that converts a matched date string like "January 1, 2020" to "01/01/2020".
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
            // Extract month, day and year from the regex groups.
            var match = args.Match;
            string monthName = match.Groups[1].Value;
            string dayStr = match.Groups[2].Value;
            string year = match.Groups[3].Value;

            if (!MonthMap.TryGetValue(monthName, out int monthNumber))
                return ReplaceAction.Skip; // Unexpected month name.

            if (!int.TryParse(dayStr, out int day))
                return ReplaceAction.Skip; // Unexpected day format.

            // Build the replacement string in MM/dd/yyyy format.
            string replacement = $"{monthNumber:D2}/{day:D2}/{year}";
            args.Replacement = replacement;
            return ReplaceAction.Replace;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Step 1: Create a sample document containing date strings.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("The project started on January 1, 2020 and ended on March 15, 2021.");
            builder.Writeln("Another milestone was on December 31, 2022.");
            string inputPath = "input.docx";
            doc.Save(inputPath);

            // Step 2: Load the document for processing.
            Document loadedDoc = new Document(inputPath);

            // Step 3: Define the regex pattern to match dates like "January 1, 2020".
            Regex dateRegex = new Regex(@"\b(January|February|March|April|May|June|July|August|September|October|November|December) (\d{1,2}), (\d{4})\b");

            // Step 4: Set up find-replace options with the custom callback.
            FindReplaceOptions options = new FindReplaceOptions();
            options.ReplacingCallback = new DateReplacer();

            // Step 5: Perform the replacement. The replacement string argument is ignored because the callback supplies it.
            int replacedCount = loadedDoc.Range.Replace(dateRegex, string.Empty, options);

            // Validate that at least one replacement occurred.
            if (replacedCount == 0)
                throw new InvalidOperationException("No date strings were replaced.");

            // Step 6: Save the modified document.
            string outputPath = "output.docx";
            loadedDoc.Save(outputPath);
        }
    }
}
