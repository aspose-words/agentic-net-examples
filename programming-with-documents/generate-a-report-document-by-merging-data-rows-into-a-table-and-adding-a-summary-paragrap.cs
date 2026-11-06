using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Sample data to merge into the table.
        var employees = new List<Employee>
        {
            new Employee { Name = "Alice Johnson", Age = 30, Salary = 75000 },
            new Employee { Name = "Bob Smith", Age = 45, Salary = 88000 },
            new Employee { Name = "Carol Davis", Age = 28, Salary = 62000 },
            new Employee { Name = "David Brown", Age = 35, Salary = 71000 }
        };

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a table with a header row.
        builder.StartTable();

        // Header cells.
        builder.InsertCell();
        builder.Write("Name");
        builder.InsertCell();
        builder.Write("Age");
        builder.InsertCell();
        builder.Write("Salary");
        builder.EndRow();

        // Populate table rows with employee data.
        foreach (var emp in employees)
        {
            builder.InsertCell();
            builder.Write(emp.Name);
            builder.InsertCell();
            builder.Write(emp.Age.ToString());
            builder.InsertCell();
            builder.Write(emp.Salary.ToString("C"));
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Add a blank line before the summary.
        builder.Writeln();

        // Compute summary values.
        int totalEmployees = employees.Count;
        double averageAge = 0;
        decimal totalSalary = 0;
        if (totalEmployees > 0)
        {
            averageAge = (double)employees.Sum(e => e.Age) / totalEmployees;
            totalSalary = employees.Sum(e => e.Salary);
        }

        // Insert a summary paragraph.
        builder.Writeln($"Total Employees: {totalEmployees}");
        builder.Writeln($"Average Age: {averageAge:F1}");
        builder.Writeln($"Total Salary: {totalSalary:C}");

        // Define output file path.
        string outputPath = "Report.docx";

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            // Optionally, reopen to ensure it is a valid document.
            Document verifyDoc = new Document(outputPath);
        }
    }

    // Simple data model for demonstration.
    private class Employee
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public decimal Salary { get; set; }
    }
}
