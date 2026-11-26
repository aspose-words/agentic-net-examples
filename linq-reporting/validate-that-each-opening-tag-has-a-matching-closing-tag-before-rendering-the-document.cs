using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingTagValidator
{
    // Sample data model
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public int Index { get; set; }
        public string Name { get; set; } = string.Empty;
    }

    public class Program
    {
        // Entry point
        public static void Main()
        {
            // Prepare sample data
            var model = new ReportModel
            {
                Items = new()
                {
                    new Item { Index = 1, Name = "Alpha" },
                    new Item { Index = 2, Name = "Beta" },
                    new Item { Index = 3, Name = "Gamma" }
                }
            };

            // Create template document programmatically
            var templatePath = "template.docx";
            CreateTemplate(templatePath);

            // Load the template for processing
            var doc = new Document(templatePath);

            // Validate LINQ Reporting tags before building the report
            if (!ValidateTags(doc, out string validationError))
            {
                Console.WriteLine($"Tag validation failed: {validationError}");
                return;
            }

            // Build the report
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated report
            var outputPath = "output.docx";
            doc.Save(outputPath);
            Console.WriteLine($"Report generated successfully: {outputPath}");
        }

        // Creates a simple template containing foreach and if tags
        private static void CreateTemplate(string path)
        {
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("LINQ Reporting Tag Validation Example");
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("Item <<[item.Index]>>: <<if [item.Index % 2 == 0]>>Even<<[item.Name]>> <</if>><<if [item.Index % 2 != 0]>>Odd<<[item.Name]>> <</if>>");
            builder.Writeln("<</foreach>>");

            doc.Save(path);
        }

        // Validates that each opening tag has a matching closing tag
        private static bool ValidateTags(Document doc, out string errorMessage)
        {
            // Extract raw text (including LINQ Reporting tags) from the document
            string docText = doc.GetText();

            // Regex to match both opening and closing tags
            // Opening tags: <<foreach ...>>, <<if ...>>, <<bookmark ...>>
            // Closing tags: <</foreach>>, <</if>>, <</bookmark>>
            var tagPattern = new Regex(@"<<\s*(?<type>foreach|if|bookmark)\b[^>]*>>|<</\s*(?<typeClose>foreach|if|bookmark)\s*>>",
                RegexOptions.Compiled | RegexOptions.IgnoreCase);

            var stack = new Stack<string>();
            foreach (Match match in tagPattern.Matches(docText))
            {
                // Determine if this is an opening or closing tag
                var opening = match.Groups["type"];
                var closing = match.Groups["typeClose"];

                if (opening.Success)
                {
                    // Push the tag type onto the stack
                    stack.Push(opening.Value.ToLowerInvariant());
                }
                else if (closing.Success)
                {
                    if (stack.Count == 0)
                    {
                        errorMessage = $"Unexpected closing tag </{closing.Value}> without matching opening tag.";
                        return false;
                    }

                    var expected = stack.Pop();
                    if (!string.Equals(expected, closing.Value, StringComparison.OrdinalIgnoreCase))
                    {
                        errorMessage = $"Mismatched tag: expected closing </{expected}> but found </{closing.Value}>.";
                        return false;
                    }
                }
            }

            if (stack.Count > 0)
            {
                errorMessage = $"Unclosed tag(s): {string.Join(", ", stack)}.";
                return false;
            }

            errorMessage = string.Empty;
            return true;
        }
    }
}
