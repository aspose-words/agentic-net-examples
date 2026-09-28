using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the temporary template and the final report.
        const string templatePath = "Template.docx";
        const string outputPath = "Report.docx";

        // 1. Create a simple template document with a table and merge fields.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Employee Report");
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Name");
        builder.InsertCell();
        builder.Write("Salary");
        builder.EndRow();

        // Row that will be repeated for each employee.
        builder.InsertCell();
        builder.InsertField("MERGEFIELD Name \\* MERGEFORMAT");
        builder.InsertCell();
        builder.InsertField("MERGEFIELD Salary \\* MERGEFORMAT");
        builder.EndRow();
        builder.EndTable();

        // Save the template to disk (required by the task to not assume external files exist).
        templateDoc.Save(templatePath);

        // 2. Generate a large data source.
        var employees = new List<Employee>();
        for (int i = 0; i < 100_000; i++)
        {
            employees.Add(new Employee { Name = $"Employee {i + 1}", Salary = 30000 + i });
        }

        // 3. Set up a cancellation token that triggers after a short delay.
        var cts = new CancellationTokenSource();
        // Cancel after 50 ms to simulate aborting a large report.
        Task.Delay(50).ContinueWith(_ => cts.Cancel());

        // 4. Prepare a LINQ query that checks for cancellation.
        IEnumerable<Employee> query = employees.Select(e =>
        {
            if (cts.Token.IsCancellationRequested)
                throw new OperationCanceledException();
            return e;
        });

        // 5. Build the report using Aspose.Words ReportingEngine.
        var engine = new ReportingEngine();

        try
        {
            // The data source object must be a visible type (public class) that exposes a property named "Employees".
            var dataSource = new ReportData { Employees = query };
            // Use the overload that accepts an output file path.
            engine.BuildReport(templateDoc, dataSource, outputPath);
            Console.WriteLine($"Report generated successfully: {outputPath}");
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Report generation was canceled.");
        }
        finally
        {
            // Clean up temporary files.
            if (File.Exists(templatePath))
                File.Delete(templatePath);
        }
    }

    // Simple POCO representing an employee.
    public class Employee
    {
        public string Name { get; set; }
        public double Salary { get; set; }
    }

    // Visible data source class required by ReportingEngine.
    public class ReportData
    {
        public IEnumerable<Employee> Employees { get; set; }
    }
}
