using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a template document in memory.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Header for the report.
        builder.Writeln("Employees Report");
        builder.Writeln();

        // Data band: iterate over the collection named 'ds'.
        // Inside the band we can refer to members of the current item
        // directly (e.g., [Name], [Age]) without specifying an object.
        builder.Writeln("<<foreach [in ds]>>");
        builder.Writeln("- <<[Name]>> (Age: <<[Age]>>)");
        builder.Writeln("<<endforeach>>");

        // -----------------------------------------------------------------
        // 2. Prepare the data source.
        // -----------------------------------------------------------------
        var dataSource = new
        {
            ds = new List<Employee>
            {
                new Employee { Name = "John Doe", Age = 30 },
                new Employee { Name = "Jane Smith", Age = 25 },
                new Employee { Name = "Bob Johnson", Age = 45 }
            }
        };

        // -----------------------------------------------------------------
        // 3. Configure the ReportingEngine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine
        {
            // Allow missing members so the engine does not throw if a member is not found.
            Options = ReportBuildOptions.AllowMissingMembers
        };
        // Text to display when a member is missing.
        engine.MissingMemberMessage = "N/A";

        // -----------------------------------------------------------------
        // 4. Build the report.
        // -----------------------------------------------------------------
        // The third parameter ('data') is the name used to reference the
        // whole data source object inside the template. Inside the foreach
        // we only need to use the collection name ('ds') and then the
        // member names of each item.
        engine.BuildReport(template, dataSource, "data");

        // -----------------------------------------------------------------
        // 5. Save the generated document.
        // -----------------------------------------------------------------
        template.Save("EmployeesReport.docx");
    }

    // Simple POCO class used as an item in the collection.
    public class Employee
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }
}
