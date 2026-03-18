using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class InTableListReport
{
    static void Main()
    {
        // Paths to the template, JSON source and the output document.
        string templatePath = @"C:\Docs\InTableListTemplate.docx";
        string jsonPath = @"C:\Docs\PeopleData.json";
        string outputPath = @"C:\Docs\PeopleReport.docx";

        // Load the Word template that contains the in‑table list syntax.
        Document doc = new Document(templatePath);

        // Create a JSON data source from the file.
        JsonDataSource jsonData = new JsonDataSource(jsonPath);

        // Build the report by merging the JSON data into the template.
        // The second parameter is the data source object,
        // the third parameter is the name used inside the template to reference the source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, jsonData, "people");

        // Save the populated document.
        doc.Save(outputPath);
    }
}
