using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;          // ReportingEngine, CsvDataLoadOptions, CsvDataSource

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create a sample CSV file.
        string csvPath = Path.Combine(outputDir, "people.csv");
        File.WriteAllLines(csvPath, new[]
        {
            "Name,Age",
            "Alice,30",
            "Bob,25"
        });

        // 2. Build a template document programmatically.
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Static section – should stay unchanged.
        builder.Writeln("=== Static Section ===");
        builder.Writeln("This paragraph must remain even if empty after processing.");

        // Start a new section that will be populated from CSV.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("=== CSV Section ===");
        // Insert LINQ Reporting tags.
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");
        // Intentionally add an empty paragraph that should be removed after the report.
        builder.Writeln();
        builder.Writeln("<</foreach>>");

        // Save the template.
        template.Save(templatePath);

        // 3. Load the template for reporting.
        Document doc = new Document(templatePath);

        // 4. Create CSV data source with headers.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        CsvDataSource csvData = new CsvDataSource(csvPath, loadOptions);

        // 5. Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        // Do NOT set RemoveEmptyParagraphs globally – we will handle it manually for the CSV section.
        engine.BuildReport(doc, csvData, "persons");

        // 6. Remove empty paragraphs only from the CSV section (the second section).
        if (doc.Sections.Count > 1)
        {
            Section csvSection = doc.Sections[1];
            List<Paragraph> emptyParagraphs = new();

            foreach (Paragraph para in csvSection.Body.Paragraphs)
            {
                // GetText includes the paragraph break; trim to check for emptiness.
                if (string.IsNullOrWhiteSpace(para.GetText()))
                    emptyParagraphs.Add(para);
            }

            foreach (Paragraph para in emptyParagraphs)
                para.Remove();
        }

        // 7. Save the final document.
        string resultPath = Path.Combine(outputDir, "Result.docx");
        doc.Save(resultPath);
    }
}
