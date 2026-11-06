using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing; // Required by the category rules
using Newtonsoft.Json; // Required by the category rules

namespace FindAndReplaceExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a sample document with custom delimiters (e.g., '|')
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Item1|Item2");
            builder.Writeln("ValueA | ValueB");
            builder.Writeln("DataX| DataY");
            builder.Writeln("Alpha |Beta|Gamma");
            // Save the initial document (optional, just for reference)
            doc.Save("input.docx");

            // Define a regular expression that matches the custom delimiter '|'
            Regex delimiterRegex = new Regex(@"\|");

            // Perform the replacement: replace each '|' with a comma ',' while preserving surrounding whitespace
            FindReplaceOptions options = new FindReplaceOptions();
            int replacedCount = doc.Range.Replace(delimiterRegex, ",", options);

            // Validate that at least one replacement occurred
            if (replacedCount == 0)
                throw new InvalidOperationException("Expected at least one delimiter replacement, but none were made.");

            // Save the modified document
            doc.Save("output.docx");

            // Optionally, write a simple JSON report of the operation (demonstrating required package usage)
            var report = new
            {
                InputFile = Path.GetFullPath("input.docx"),
                OutputFile = Path.GetFullPath("output.docx"),
                DelimiterReplaced = "|",
                Replacement = ",",
                ReplacementsMade = replacedCount
            };
            string jsonReport = JsonConvert.SerializeObject(report, Formatting.Indented);
            File.WriteAllText("report.json", jsonReport);
        }
    }
}
