using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportData
{
    public string Title { get; set; }
    public string Chinese { get; set; }
    public string Arabic { get; set; }
    public string Russian { get; set; }
}

class MultilingualReportGenerator
{
    static void Main()
    {
        // 1. Create a simple template document with placeholders that reference object fields.
        //    The placeholders use the Reporting Engine syntax: <<[data.FieldName]>>.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Report Title: <<[data.Title]>>");
        builder.Writeln("Chinese (简体中文): <<[data.Chinese]>>");
        builder.Writeln("Arabic (العربية): <<[data.Arabic]>>");
        builder.Writeln("Russian (Русский): <<[data.Russian]>>");

        // Optional: save the template for inspection.
        template.Save("Template.docx");

        // 2. Prepare the data source as a public class instance containing multilingual fields.
        var data = new ReportData
        {
            Title   = "Multilingual Report",
            Chinese = "你好，世界！",
            Arabic  = "مرحبا بالعالم!",
            Russian = "Привет, мир!"
        };

        // 3. Populate the template with data using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        // The third argument ("data") is the name used inside the template to reference the source object.
        engine.BuildReport(template, data, "data");

        // 4. Save the final report as a DOCX file. Unicode characters are preserved automatically.
        template.Save("MultilingualReport.docx");
    }
}
