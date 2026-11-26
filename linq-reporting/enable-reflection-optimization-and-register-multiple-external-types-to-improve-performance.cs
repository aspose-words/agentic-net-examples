using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for any encoding needs.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create sample hierarchical data.
        var company = new Company
        {
            Name = "Tech Solutions Inc.",
            Departments = new List<Department>
            {
                new Department
                {
                    Name = "Research & Development",
                    Location = "Building A",
                    Employees = new List<Employee>
                    {
                        new Employee
                        {
                            FirstName = "Alice",
                            LastName = "Johnson",
                            Position = "Senior Engineer",
                            Address = new Address
                            {
                                Street = "123 Innovation Way",
                                City = "Techville"
                            }
                        },
                        new Employee
                        {
                            FirstName = "Bob",
                            LastName = "Smith",
                            Position = "Junior Engineer",
                            Address = new Address
                            {
                                Street = "456 Discovery Rd",
                                City = "Techville"
                            }
                        }
                    }
                },
                new Department
                {
                    Name = "Human Resources",
                    Location = "Building B",
                    Employees = new List<Employee>
                    {
                        new Employee
                        {
                            FirstName = "Carol",
                            LastName = "White",
                            Position = "HR Manager",
                            Address = new Address
                            {
                                Street = "789 People St",
                                City = "Admin City"
                            }
                        }
                    }
                }
            }
        };

        // Create a template document programmatically.
        const string templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Company: <<[company.Name]>>");
        builder.Writeln();

        // Departments loop
        builder.Writeln("<<foreach [dept in company.Departments]>>");
        builder.Writeln("Department: <<[dept.Name]>>");
        builder.Writeln("Location: <<[dept.Location]>>");
        builder.Writeln("Employees:");
        // Employees loop
        builder.Writeln("<<foreach [emp in dept.Employees]>>");
        builder.Writeln("- <<[emp.FirstName]>> <<[emp.LastName]>> (<<[emp.Position]>>)");
        builder.Writeln("  Address: <<[emp.Address.Street]>>, <<[emp.Address.City]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);

        // Enable reflection optimization.
        ReportingEngine.UseReflectionOptimization = true;

        // Create the reporting engine.
        var engine = new ReportingEngine();

        // Build the report.
        bool success = engine.BuildReport(reportDoc, company, "company");

        // Save the generated report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Output saved to '{outputPath}'.");
    }
}

// Data model classes.
public class Company
{
    public string Name { get; set; } = string.Empty;
    public List<Department> Departments { get; set; } = new();
}

public class Department
{
    public string Name { get; set; } = string.Empty;
    public string Location { get; set; } = string.Empty;
    public List<Employee> Employees { get; set; } = new();
}

public class Employee
{
    public string FirstName { get; set; } = string.Empty;
    public string LastName { get; set; } = string.Empty;
    public string Position { get; set; } = string.Empty;
    public Address Address { get; set; } = new();
}

public class Address
{
    public string Street { get; set; } = string.Empty;
    public string City { get; set; } = string.Empty;
}
