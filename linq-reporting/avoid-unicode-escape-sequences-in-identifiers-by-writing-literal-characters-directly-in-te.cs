using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for any required encodings.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // -------------------- Create template --------------------
        // The template contains a LINQ Reporting tag that references a Unicode property directly.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Customer name: <<[model.名字]>>");

        // Save the template to disk (demonstrates the create‑save lifecycle).
        const string templatePath = "Template.docx";
        doc.Save(templatePath);

        // -------------------- Load template --------------------
        Document template = new Document(templatePath);

        // -------------------- Prepare data model --------------------
        // The model uses a literal Unicode identifier for the property name.
        var model = new ReportModel
        {
            名字 = "张三"
        };

        // -------------------- Build report --------------------
        ReportingEngine engine = new ReportingEngine();
        // The root object name in the template is "model", so we pass it explicitly.
        engine.BuildReport(template, model, "model");

        // -------------------- Save result --------------------
        const string outputPath = "Report.docx";
        template.Save(outputPath);
    }
}

// Data model with a Unicode property name written directly (no escape sequences).
public class ReportModel
{
    // Initialize to avoid nullable warnings.
    public string 名字 { get; set; } = string.Empty;
}
