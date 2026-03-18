using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class ConvertDateFormat
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Sample text containing dates in MM-DD-YYYY format.
        builder.Writeln("The project started on 03-15-2022 and ended on 12-31-2022.");
        builder.Writeln("Another date: 07-04-2021.");

        // Regular expression to capture month, day and year.
        Regex datePattern = new Regex(@"(\d{2})-(\d{2})-(\d{4})");

        // FindReplaceOptions must enable substitutions to use $1, $2, $3 in the replacement string.
        FindReplaceOptions options = new FindReplaceOptions
        {
            UseSubstitutions = true
        };

        // Replace MM-DD-YYYY with YYYY-MM-DD.
        doc.Range.Replace(datePattern, "$3-$1-$2", options);

        // Save the modified document.
        doc.Save("ConvertedDates.docx");
    }
}
