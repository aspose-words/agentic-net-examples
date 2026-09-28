using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingTagValidation
{
    // Sample data model
    public class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }

    public class ReportModel
    {
        public List<Person> Persons { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some encodings)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for template and output
            string templatePath = "Template.docx";
            string outputPath = "Report.docx";

            // 1. Create template document with LINQ Reporting tags
            CreateTemplate(templatePath);

            // 2. Load the template
            Document doc = new Document(templatePath);

            // 3. Validate tags before building the report
            bool tagsValid = ValidateReportingTags(doc);
            Console.WriteLine($"Tag validation result: {(tagsValid ? "Success" : "Failure")}");

            if (!tagsValid)
            {
                // Abort if validation fails
                return;
            }

            // 4. Prepare data
            ReportModel model = new()
            {
                Persons = new()
                {
                    new Person { Name = "Alice", Age = 30 },
                    new Person { Name = "Bob", Age = 25 },
                    new Person { Name = "Charlie", Age = 35 }
                }
            };

            // 5. Build report
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // 6. Save the generated report
            doc.Save(outputPath);
            Console.WriteLine($"Report generated: {outputPath}");
        }

        private static void CreateTemplate(string path)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Title
            builder.Writeln("People Report");
            builder.Writeln();

            // Begin foreach loop
            builder.Writeln("<<foreach [p in Persons]>>");
            // Conditional formatting: highlight age > 30
            builder.Writeln("<<if [p.Age > 30]>>");
            builder.Writeln("<<backColor [\"LightYellow\"]>><<[p.Name]>> (Age: <<[p.Age]>>) <</backColor>>");
            builder.Writeln("<</if>>");
            builder.Writeln("<<if [p.Age <= 30]>>");
            builder.Writeln("<<[p.Name]>> (Age: <<[p.Age]>>)");
            builder.Writeln("<</if>>");
            builder.Writeln("<</foreach>>");

            // Save template
            doc.Save(path);
        }

        private static bool ValidateReportingTags(Document doc)
        {
            // Retrieve full document text (includes LINQ Reporting tags)
            string text = doc.GetText();

            // Regex patterns for opening and closing tags
            string openingPattern = @"<<\s*(foreach|if|bookmark|textColor|backColor)\b[^>]*>>";
            string closingPattern = @"<</\s*(foreach|if|bookmark|textColor|backColor)\s*>>";

            // Find all tags in order
            var tagRegex = new Regex($"{openingPattern}|{closingPattern}", RegexOptions.Compiled);
            var matches = tagRegex.Matches(text);

            Stack<string> stack = new();

            foreach (Match match in matches)
            {
                string tag = match.Value;

                // Opening tag
                var openMatch = Regex.Match(tag, @"<<\s*(foreach|if|bookmark|textColor|backColor)\b", RegexOptions.IgnoreCase);
                if (openMatch.Success)
                {
                    string tagName = openMatch.Groups[1].Value.ToLowerInvariant();
                    stack.Push(tagName);
                    continue;
                }

                // Closing tag
                var closeMatch = Regex.Match(tag, @"<</\s*(foreach|if|bookmark|textColor|backColor)\s*>>", RegexOptions.IgnoreCase);
                if (closeMatch.Success)
                {
                    string tagName = closeMatch.Groups[1].Value.ToLowerInvariant();
                    if (stack.Count == 0)
                    {
                        Console.WriteLine($"Unmatched closing tag: {tag}");
                        return false;
                    }

                    string expected = stack.Pop();
                    if (expected != tagName)
                    {
                        Console.WriteLine($"Mismatched tag. Expected closing for '{expected}' but found '{tagName}'.");
                        return false;
                    }
                }
            }

            if (stack.Count > 0)
            {
                Console.WriteLine($"Unmatched opening tag(s): {string.Join(", ", stack)}");
                return false;
            }

            return true;
        }
    }
}
