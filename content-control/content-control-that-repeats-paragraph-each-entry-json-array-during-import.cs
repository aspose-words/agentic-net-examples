using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Reporting;

class RepeatingSectionFromJson
{
    static void Main()
    {
        // 1. Create a template document that contains a repeating section content control.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert a table header (optional – the repeating section can be inside a table or a plain paragraph).
        builder.Writeln("People List:");

        // Create the repeating section (block‑level content control).
        StructuredDocumentTag repeatingSection = new StructuredDocumentTag(template,
            SdtType.RepeatingSection, MarkupLevel.Block);
        builder.InsertNode(repeatingSection); // place the control in the document.

        // Inside the repeating section we need a repeating section item.
        StructuredDocumentTag repeatingItem = new StructuredDocumentTag(template,
            SdtType.RepeatingSectionItem, MarkupLevel.Block);
        repeatingSection.AppendChild(repeatingItem);

        // The item will contain a single paragraph that will be repeated for each JSON array entry.
        Paragraph para = new Paragraph(template);
        // Use ReportingEngine syntax to refer to a field inside the JSON object.
        // Assuming each array element has a property called "Name".
        para.AppendChild(new Run(template, "<<[person.Name]>>"));
        repeatingItem.AppendChild(para);

        // Save the template – this is the "create" rule.
        const string templatePath = "RepeatingSectionTemplate.docx";
        template.Save(templatePath);

        // 2. Load the JSON data.
        // Example JSON:
        // [
        //   { "Name": "Alice" },
        //   { "Name": "Bob" },
        //   { "Name": "Charlie" }
        // ]
        const string jsonPath = "people.json";

        // Create a JsonDataSource that reads the file.
        JsonDataSource jsonData = new JsonDataSource(jsonPath);

        // 3. Build the report – the ReportingEngine will repeat the paragraph for each array element.
        // The data source name ("person") must match the name used in the template field.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, jsonData, "person");

        // 4. Save the final document – this is the "save" rule.
        const string outputPath = "RepeatingSectionResult.docx";
        template.Save(outputPath);
    }
}
