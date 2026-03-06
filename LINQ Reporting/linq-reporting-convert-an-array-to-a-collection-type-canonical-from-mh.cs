using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

class LinqReportingArrayToCollection
{
    static void Main()
    {
        // Load the template document that is stored in MHTML format.
        // The Document constructor automatically detects the format from the file extension.
        Document template = new Document(@"MyDir\Template.mhtml");

        // Original data source is an array of strings.
        string[] nameArray = new string[] { "Alice", "Bob", "Charlie" };

        // Convert the array to a canonical collection type (List<string>) using LINQ.
        // ReportingEngine works best with collection types that implement IEnumerable<T>.
        List<string> nameList = nameArray.ToList();

        // Create the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Build the report using the template and the collection as a data source.
        // The third parameter is the name by which the data source will be referenced in the template.
        engine.BuildReport(template, nameList, "Names");

        // Save the populated document to DOCX format.
        template.Save(@"ArtifactsDir\Report.docx", SaveFormat.Docx);
    }
}
