using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Replacing;

public class Person
{
    public string Name { get; set; } = "John Doe";
    public int Age { get; set; } = 30;
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        string templatePath = "Template.docx";
        string reportPath = "Report.docx";

        // 1. Create a template document with custom tag delimiters [[ and ]].
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        // Use custom delimiters that will later be converted to proper LINQ Reporting tags.
        builder.Writeln("Hello [[person.Name]], Age [[person.Age]]");
        templateDoc.Save(templatePath);

        // 2. Load the template back from disk.
        Document doc = new Document(templatePath);

        // 3. Replace custom delimiters with the default LINQ Reporting delimiters <<[ and ]>>.
        // This conversion yields valid tags like <<[person.Name]>>.
        FindReplaceOptions replaceOptions = new FindReplaceOptions();
        doc.Range.Replace("[[", "<<[", replaceOptions);
        doc.Range.Replace("]]", "]>>", replaceOptions);

        // 4. Prepare the data source.
        Person person = new Person();

        // 5. Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, person, "person");

        // 6. Save the generated report.
        doc.Save(reportPath);
    }
}
