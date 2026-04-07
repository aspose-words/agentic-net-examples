using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string outputPath = "Report.docx";

        // -------------------------------------------------
        // 1. Create the template document programmatically.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // ----- Static section (no selective removal) -----
        builder.Writeln("=== Static Section ===");
        // This tag does NOT contain an exclamation mark, therefore empty paragraphs here will be kept.
        builder.Writeln("<<[staticText]>>");

        // Insert a section break to separate the static part from the JSON‑driven part.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // ----- JSON‑driven section (selective removal) -----
        builder.Writeln("=== JSON Section ===");
        // The foreach loop iterates over the JSON array "items".
        builder.Writeln("<<foreach [item in items]>>");
        // Title is always written.
        builder.Writeln("Title: <<[item.Title]>>");
        // Content tag – empty paragraphs produced by this tag will be removed because the engine
        // is configured to remove empty paragraphs globally. The tag itself does not need an
        // exclamation mark; the engine will handle removal.
        builder.Writeln("Content: <<[item.Content]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template back before building the report.
        // -------------------------------------------------
        Document doc = new Document(templatePath);

        // -------------------------------------------------
        // 3. Prepare JSON data source.
        // -------------------------------------------------
        var json = @"
        {
            ""items"": [
                { ""Title"": ""First"",  ""Content"": ""Hello World!"" },
                { ""Title"": ""Second"", ""Content"": """" },
                { ""Title"": ""Third"",  ""Content"": ""Aspose.Words"" }
            ],
            ""staticText"": ""This is static text.""
        }";

        // The ReportingEngine works with a JsonDataSource instance.
        using (MemoryStream jsonStream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(json)))
        {
            JsonDataSource jsonData = new JsonDataSource(jsonStream);

            // -------------------------------------------------
            // 4. Build the report.
            // -------------------------------------------------
            ReportingEngine engine = new ReportingEngine
            {
                // Enable removal of empty paragraphs.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // The root object name is optional; we pass an empty string because the template
            // references members directly (e.g., <<[staticText]>> and <<foreach [item in items]>>).
            engine.BuildReport(doc, jsonData, "");

            // -------------------------------------------------
            // 5. Save the final document.
            // -------------------------------------------------
            doc.Save(outputPath);
        }
    }
}
