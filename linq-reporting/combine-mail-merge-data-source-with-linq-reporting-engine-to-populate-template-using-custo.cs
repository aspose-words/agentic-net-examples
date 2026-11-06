using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a simple template with a mail‑merge field and a LINQ Reporting tag.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Mail‑merge field $customer$
        builder.Writeln("Mail‑Merge Customer: $customer$");
        // LINQ Reporting tag that will be filled from the same data source.
        builder.Writeln("LINQ Reporting Customer: <<[model.Customer]>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template back for processing.
        Document doc = new Document(templatePath);

        // Prepare the data model.
        ReportModel model = new()
        {
            Customer = "John Doe"
        };

        // Execute mail‑merge to replace $customer$ fields.
        doc.MailMerge.Execute(new[] { "customer" }, new object[] { model.Customer });

        // Build the LINQ Reporting part.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the final report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}

// Public data model used by both mail‑merge and LINQ Reporting.
public class ReportModel
{
    public string Customer { get; set; } = string.Empty;
}
