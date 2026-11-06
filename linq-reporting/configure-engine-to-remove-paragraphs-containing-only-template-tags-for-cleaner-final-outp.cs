using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare template file.
        const string templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Normal paragraph with a tag.
        builder.Writeln("Hello, <<[person.Name]>>!");

        // Paragraph that contains only a conditional tag.
        builder.Writeln("<<if [person.Age > 30]>>Senior<</if>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Sample data model.
        var person = new Person
        {
            Name = "John",
            Age = 25
        };

        // Configure the reporting engine to remove empty paragraphs.
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.RemoveEmptyParagraphs
        };

        // Build the report.
        engine.BuildReport(doc, person, "person");

        // Save the final document.
        const string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}

// Data model class.
public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}
