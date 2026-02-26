using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsReportingDemo
{
    // Simple data model representing a topic.
    public class Topic
    {
        // Made nullable to satisfy the non‑nullable warnings when using the default constructor.
        public string? Name { get; set; }
        public string? Title { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Prepare sample data.
            List<Topic> topics = new List<Topic>
            {
                new Topic { Name = "Intro",   Title = "Introduction to Reporting" },
                new Topic { Name = "Setup",   Title = "Setting up Aspose.Words" },
                new Topic { Name = "Example", Title = "Dynamic Bookmark Example" }
            };

            // 2. Create a DOT (template) document in memory.
            //    The template uses ReportingEngine syntax:
            //    <<foreach [topic]>> – iterate over the collection named "topic".
            //    <<bookmark [topic.Name]>> – start a bookmark whose name comes from the current item.
            //    <<[topic.Title]>> – insert the title text.
            //    <<endbookmark>> – close the bookmark.
            //    <<endfor>> – end the loop.
            Document template = new Document();                     // create a blank document
            DocumentBuilder builder = new DocumentBuilder(template); // helper to add content

            builder.Writeln("Table of Contents:");
            builder.Writeln(); // empty line

            // Begin the foreach loop.
            builder.Writeln("<<foreach [topic]>>");

            // Insert a bookmark start. The ReportingEngine will replace [topic.Name] with the actual name.
            builder.Writeln("<<bookmark [topic.Name]>>");

            // Insert the title that will be inside the bookmark.
            builder.Writeln("<<[topic.Title]>>");

            // Close the bookmark.
            builder.Writeln("<<endbookmark>>");

            // End the foreach loop.
            builder.Writeln("<<endfor>>");

            // Save the template to a DOT file (the extension does not affect processing, but mimics a template).
            string templatePath = "Template.dot";
            template.Save(templatePath);

            // 3. Load the template document (simulating a real‑world scenario where the template is stored on disk).
            Document loadedTemplate = new Document(templatePath);

            // 4. Build the report using ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // The overload expects a single data‑source name, not an array.
            engine.BuildReport(loadedTemplate, topics, "topic");

            // 5. Save the generated report.
            string outputPath = "ReportWithBookmarks.docx";
            loadedTemplate.Save(outputPath);

            Console.WriteLine($"Report generated successfully: {outputPath}");
        }
    }
}
