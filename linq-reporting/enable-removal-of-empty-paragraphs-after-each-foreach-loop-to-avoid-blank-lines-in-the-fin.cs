using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Register code page provider for possible data sources.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for the template and the final report.
        string templatePath = "Template.docx";
        string outputPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title.
        builder.Writeln("Customer List");
        builder.Writeln();

        // Begin foreach loop over Persons collection.
        builder.Writeln("<<foreach [p in Persons]>>");
        // Paragraph that may become empty if data is missing.
        builder.Writeln("<<[p.Name]>> - <<[p.Age]>>");
        // Add an empty paragraph inside the loop to simulate potential blank lines.
        builder.Writeln();
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and prepare data.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // Sample data model.
        ReportModel model = new()
        {
            Persons = new()
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob", Age = 25 },
                // An entry with empty values to test removal of blank lines.
                new Person { Name = "", Age = 0 }
            }
        };

        // -----------------------------------------------------------------
        // 3. Build the report using LINQ Reporting engine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // -----------------------------------------------------------------
        // 4. Remove empty paragraphs that may have been left after foreach.
        // -----------------------------------------------------------------
        RemoveEmptyParagraphs(doc);

        // -----------------------------------------------------------------
        // 5. Save the final document.
        // -----------------------------------------------------------------
        doc.Save(outputPath);
    }

    // Removes paragraphs that are blank (contain no visible text).
    private static void RemoveEmptyParagraphs(Document doc)
    {
        // Collect paragraphs to remove to avoid modifying the collection while iterating.
        List<Paragraph> emptyParagraphs = new();

        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            // A paragraph is considered empty if its text, after trimming, is empty.
            if (string.IsNullOrWhiteSpace(para.GetText()))
            {
                emptyParagraphs.Add(para);
            }
        }

        foreach (Paragraph para in emptyParagraphs)
        {
            para.Remove();
        }
    }
}

// ---------------------------------------------------------------------
// Data model classes.
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}
