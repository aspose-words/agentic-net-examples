using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a template document with a repeating section content control.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        Body body = templateDoc.FirstSection.Body;

        // Repeating section (block level).
        StructuredDocumentTag repeatingSection = new StructuredDocumentTag(templateDoc, SdtType.RepeatingSection, MarkupLevel.Block);
        body.AppendChild(repeatingSection);

        // Repeating section item (block level).
        StructuredDocumentTag repeatingItem = new StructuredDocumentTag(templateDoc, SdtType.RepeatingSectionItem, MarkupLevel.Block);
        repeatingSection.AppendChild(repeatingItem);

        // Paragraph inside the repeating item.
        Paragraph paragraph = new Paragraph(templateDoc);
        repeatingItem.AppendChild(paragraph);

        // Plain text content control that will display the "Name" field.
        StructuredDocumentTag nameControl = new StructuredDocumentTag(templateDoc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "Name",
            Tag = "Name"
        };
        paragraph.AppendChild(nameControl);

        // Save the template to disk.
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // Prepare JSON data representing an array of objects with a "Name" property.
        string json = @"[
            { ""Name"": ""Alice"" },
            { ""Name"": ""Bob"" },
            { ""Name"": ""Charlie"" }
        ]";
        const string jsonPath = "data.json";
        File.WriteAllText(jsonPath, json);

        // Load the template and populate it using the ReportingEngine with the JSON data source.
        Document reportDoc = new Document(templatePath);
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);
        ReportingEngine engine = new ReportingEngine();

        // The data source name "persons" is used in the template; the engine will repeat the
        // repeating section for each element in the JSON array.
        engine.BuildReport(reportDoc, jsonDataSource, "persons");

        // Save the resulting document.
        const string resultPath = "Result.docx";
        reportDoc.Save(resultPath);
    }
}
