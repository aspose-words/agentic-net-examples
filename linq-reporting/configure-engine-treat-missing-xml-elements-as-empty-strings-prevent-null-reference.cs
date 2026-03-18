using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the template document that contains LINQ Reporting Engine tags.
        Document doc = new Document("Template.docx");

        // Ensure that a root object is always generated for the XML data source.
        // This prevents the engine from skipping the root when it has no attributes
        // and all child elements share the same name.
        XmlDataLoadOptions xmlLoadOptions = new XmlDataLoadOptions
        {
            AlwaysGenerateRootObject = true
        };

        // Create an XML data source using the options defined above.
        // The second argument is the name of the root element used in the template.
        XmlDataSource dataSource = new XmlDataSource("Data.xml", xmlLoadOptions);

        // Configure the reporting engine:
        // - AllowMissingMembers tells the engine to treat missing members as null literals.
        // - MissingMemberMessage set to an empty string makes those nulls render as empty strings.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers,
            MissingMemberMessage = string.Empty
        };

        // Build the report. The third parameter is the name of the data source root object
        // as referenced in the template (e.g., <<[persons.name]>> would use "persons").
        engine.BuildReport(doc, dataSource, "persons");

        // Save the generated document.
        doc.Save("Result.docx");
    }
}
