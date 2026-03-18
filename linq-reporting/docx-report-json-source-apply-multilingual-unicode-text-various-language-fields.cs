using System;
using Aspose.Words;
using Aspose.Words.Reporting; // JsonDataSource resides in this namespace

class MultilingualReportGenerator
{
    static void Main()
    {
        // 1. Create a simple template document with placeholders that reference JSON fields.
        //    The placeholders use the Reporting Engine syntax: <<[data.FieldName]>>.
        Document template = new Document();                     // create blank document
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Report Title: <<[data.Title]>>");
        builder.Writeln("Chinese (简体中文): <<[data.Chinese]>>");
        builder.Writeln("Arabic (العربية): <<[data.Arabic]>>");
        builder.Writeln("Russian (Русский): <<[data.Russian]>>");

        // Optional: save the template for inspection.
        template.Save("Template.docx");                         // save template

        // 2. Load JSON data. The JSON file should contain the multilingual fields referenced above.
        //    Example JSON (data.json):
        //    {
        //        "Title": "Multilingual Report",
        //        "Chinese": "你好，世界！",
        //        "Arabic": "مرحبا بالعالم!",
        //        "Russian": "Привет, мир!"
        //    }
        string jsonPath = "data.json";
        JsonDataSource jsonSource = new JsonDataSource(jsonPath); // create JSON data source

        // 3. Populate the template with data using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();          // instantiate engine
        // The third argument ("data") is the name used inside the template to reference the source object.
        engine.BuildReport(template, jsonSource, "data");

        // 4. Save the final report as a DOCX file. Unicode characters are preserved automatically.
        template.Save("MultilingualReport.docx");               // save populated report
    }
}
