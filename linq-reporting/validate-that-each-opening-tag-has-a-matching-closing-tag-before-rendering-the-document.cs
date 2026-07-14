using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model used by the LINQ Reporting template.
    public class ReportModel
    {
        // Sample scalar property.
        public string Name { get; set; } = "World";

        // Collection used by a foreach tag.
        public List<string> Items { get; set; } = new();
    }

    public class Program
    {
        // Entry point of the console application.
        public static void Main()
        {
            // Paths for the template and the final report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -------------------------------------------------
            // 1. Create the LINQ Reporting template programmatically.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Simple scalar tag.
            builder.Writeln("Hello <<[model.Name]>>!");

            // Foreach tag that iterates over Items collection.
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("- <<[item]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk (required before loading).
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template back from the file system.
            // -------------------------------------------------
            Document doc = new Document(templatePath);

            // -------------------------------------------------
            // 3. Validate that every opening tag has a matching closing tag.
            // -------------------------------------------------
            if (!ValidateTags(doc))
            {
                Console.WriteLine("Tag validation failed. The document will not be rendered.");
                return;
            }

            // -------------------------------------------------
            // 4. Prepare the data source.
            // -------------------------------------------------
            ReportModel model = new ReportModel
            {
                Name = "Aspose.Words",
                Items = new List<string> { "Item A", "Item B", "Item C" }
            };

            // -------------------------------------------------
            // 5. Build the report using the ReportingEngine.
            // -------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // -------------------------------------------------
            // 6. Save the generated report.
            // -------------------------------------------------
            doc.Save(reportPath);
            Console.WriteLine($"Report generated successfully: {Path.GetFullPath(reportPath)}");
        }

        // Validates that each opening tag (e.g., <<foreach ...>>) has a corresponding closing tag (e.g., <</foreach>>).
        private static bool ValidateTags(Document document)
        {
            // Retrieve the full text of the document, which contains the LINQ Reporting tags.
            string text = document.GetText();

            // Regular expression to capture all tags of the form <<...>>.
            Regex tagRegex = new Regex(@"<<([^>]+)>>", RegexOptions.Compiled);
            MatchCollection matches = tagRegex.Matches(text);

            // Stack to keep track of opening tags.
            Stack<string> tagStack = new Stack<string>();

            foreach (Match match in matches)
            {
                string tagContent = match.Groups[1].Value.Trim();

                // Closing tags start with a forward slash, e.g., /foreach, /if, /bookmark.
                if (tagContent.StartsWith("/"))
                {
                    string closingTagName = tagContent.Substring(1).Split(' ')[0];

                    if (tagStack.Count == 0)
                    {
                        Console.WriteLine($"Unexpected closing tag: </{closingTagName}>");
                        return false;
                    }

                    string openingTagName = tagStack.Pop();
                    if (!string.Equals(openingTagName, closingTagName, StringComparison.OrdinalIgnoreCase))
                    {
                        Console.WriteLine($"Mismatched tag pair: opened with <{openingTagName}> but closed with </{closingTagName}>");
                        return false;
                    }
                }
                else
                {
                    // Opening tag – extract the tag name (first token before any space or bracket).
                    string openingTagName = tagContent.Split(new[] { ' ', '[' }, StringSplitOptions.RemoveEmptyEntries)[0];
                    if (IsTagRequiringClosure(openingTagName))
                    {
                        tagStack.Push(openingTagName);
                    }
                }
            }

            if (tagStack.Count > 0)
            {
                Console.WriteLine($"Unclosed tag detected: <{tagStack.Peek()}>");
                return false;
            }

            // All tags are properly balanced.
            return true;
        }

        // Determines whether a tag requires an explicit closing tag.
        private static bool IsTagRequiringClosure(string tagName)
        {
            // Tags that have explicit closing forms in LINQ Reporting.
            var tagsRequiringClosure = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "foreach",
                "if",
                "bookmark",
                "cellmerge",
                "restartnum"
            };

            return tagsRequiringClosure.Contains(tagName);
        }
    }
}
