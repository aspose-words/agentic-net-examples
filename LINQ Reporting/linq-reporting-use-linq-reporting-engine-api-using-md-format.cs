using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Markdown template containing LINQ Reporting Engine tags
        string mdTemplate = @"
# Employee Report

<<foreach [person in persons]>>
- **Name:** <<[person.Name]>>
- **Age:** <<[person.Age]>>
- **Department:** <<[person.Department]>>

<<endforeach>>
";

        // Load the markdown string into an Aspose.Words Document
        Document doc;
        using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(mdTemplate)))
        {
            doc = new Document(ms);
        }

        // Prepare the data source for the report
        var data = new
        {
            persons = new[]
            {
                new { Name = "John Doe", Age = 30, Department = "Sales" },
                new { Name = "Jane Smith", Age = 28, Department = "HR" }
            }
        };

        // Build the report using the LINQ Reporting Engine
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, data, ""); // Empty name means we reference members directly

        // Save the generated document
        doc.Save("EmployeeReport.docx");
    }
}
