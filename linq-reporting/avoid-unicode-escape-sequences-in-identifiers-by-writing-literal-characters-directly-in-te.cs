using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare output folder
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // File paths for the template and the generated report
        string templatePath = Path.Combine(outputDir, "Template.docx");
        string resultPath = Path.Combine(outputDir, "Result.docx");

        // -------------------------------------------------
        // 1. Create a template document containing a LINQ Reporting tag.
        //    The tag references a property whose name includes literal Unicode characters.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("顧客情報:");                     // Japanese header text
        builder.Writeln("<<[model.名前]>>");            // Directly use the Unicode identifier in the tag
        templateDoc.Save(templatePath);                 // Save the template before building the report

        // -------------------------------------------------
        // 2. Load the saved template for reporting.
        // -------------------------------------------------
        Document doc = new Document(templatePath);

        // -------------------------------------------------
        // 3. Create a data model with a Unicode property name.
        // -------------------------------------------------
        var model = new CustomerModel { 名前 = "山田太郎" };

        // -------------------------------------------------
        // 4. Build the report using the ReportingEngine.
        // -------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // -------------------------------------------------
        // 5. Save the generated report.
        // -------------------------------------------------
        doc.Save(resultPath);
    }
}

// Data model class with a property that uses literal Unicode characters in its name.
public class CustomerModel
{
    // Property name contains Japanese characters directly (no escape sequences).
    public string 名前 { get; set; } = string.Empty;
}
