using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingTagValidation
{
    // Sample data model
    public class Model
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Name { get; set; } = string.Empty;
        public bool IsActive { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the temporary template and final report
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -------------------------------------------------
            // 1. Create the template document programmatically
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Add a foreach block
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("Name: <<[item.Name]>>");
            // Add an if block inside the foreach
            builder.Writeln("<<if [item.IsActive]>>Status: Active<</if>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk (required before BuildReport)
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template back from disk
            // -------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // -------------------------------------------------
            // 3. Validate that every opening tag has a matching closing tag
            // -------------------------------------------------
            if (!ValidateTags(loadedTemplate))
            {
                throw new InvalidOperationException("Tag validation failed: mismatched opening/closing tags.");
            }

            // -------------------------------------------------
            // 4. Prepare sample data
            // -------------------------------------------------
            Model model = new Model
            {
                Items = new List<Item>
                {
                    new Item { Name = "Alice", IsActive = true },
                    new Item { Name = "Bob",   IsActive = false },
                    new Item { Name = "Carol", IsActive = true }
                }
            };

            // -------------------------------------------------
            // 5. Build the report using the LINQ Reporting engine
            // -------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(loadedTemplate, model, "model");

            // -------------------------------------------------
            // 6. Save the generated report
            // -------------------------------------------------
            loadedTemplate.Save(reportPath);
        }

        // Simple validation that counts opening and closing tags for supported constructs
        private static bool ValidateTags(Document doc)
        {
            string text = doc.GetText();

            // Define tag pairs to check
            var tagPairs = new Dictionary<string, (string Open, string Close)>
            {
                { "foreach", ("<<foreach", "<</foreach>>") },
                { "if",      ("<<if",      "<</if>>") },
                { "bookmark",("<<bookmark","<</bookmark>>") },
                { "textColor",("<<textColor", "<</textColor>>") },
                { "backColor",("<<backColor", "<</backColor>>") },
                { "cellMerge",("<<cellMerge", "<</cellMerge>>") },
                { "restartNum",("<<restartNum", "<</restartNum>>") }
                // Add more pairs as needed
            };

            foreach (var pair in tagPairs.Values)
            {
                int openCount = CountOccurrences(text, pair.Open);
                int closeCount = CountOccurrences(text, pair.Close);
                if (openCount != closeCount)
                {
                    // Mismatch found
                    return false;
                }
            }

            // All checked tags are balanced
            return true;
        }

        // Helper to count non‑overlapping occurrences of a substring
        private static int CountOccurrences(string source, string substring)
        {
            if (string.IsNullOrEmpty(substring))
                return 0;

            int count = 0;
            int index = 0;
            while ((index = source.IndexOf(substring, index, StringComparison.Ordinal)) != -1)
            {
                count++;
                index += substring.Length;
            }
            return count;
        }
    }
}
