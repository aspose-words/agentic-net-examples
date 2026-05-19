using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Sample static helper class that will be used inside the template.
    public static class CustomHelper
    {
        public static string Upper(string value) => value?.ToUpperInvariant() ?? string.Empty;
    }

    // Root wrapper class – the template will reference this object as "model".
    public class ReportModel
    {
        public Company Company { get; set; } = new();
    }

    public class Company
    {
        public string Name { get; set; } = string.Empty;
        public List<Department> Departments { get; set; } = new();
    }

    public class Department
    {
        public string DeptName { get; set; } = string.Empty;
        public List<Employee> Employees { get; set; } = new();
    }

    public class Employee
    {
        public string FullName { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Enable reflection optimization for the reporting engine.
            ReportingEngine.UseReflectionOptimization = true;

            // Create a new document and build the LINQ Reporting template.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Simple header.
            builder.Writeln("Company Report");
            builder.Writeln("Company: <<[model.Company.Name]>>");
            builder.Writeln();

            // Iterate over departments.
            builder.Writeln("<<foreach [dept in model.Company.Departments]>>");
            builder.Writeln("Department: <<[dept.DeptName]>>");
            builder.Writeln("Employees:");

            // Iterate over employees inside each department.
            builder.Writeln("<<foreach [emp in dept.Employees]>>");
            // Use a registered static type (CustomHelper) to transform the employee name.
            builder.Writeln("- <<[CustomHelper.Upper(emp.FullName)]>> (Age: <<[emp.Age]>>)");
            builder.Writeln("<</foreach>>"); // End employee foreach.
            builder.Writeln("<</foreach>>"); // End department foreach.

            // Initialize the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Register external types that can be accessed from the template.
            engine.KnownTypes.Add(typeof(CustomHelper));
            engine.KnownTypes.Add(typeof(System.Math));

            // Build sample hierarchical data.
            ReportModel model = CreateSampleData();

            // Build the report using the model and the root name "model".
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            doc.Save("ReportOutput.docx");
        }

        // Generates realistic sample data for the report.
        private static ReportModel CreateSampleData()
        {
            var model = new ReportModel
            {
                Company = new Company
                {
                    Name = "Tech Solutions Ltd.",
                    Departments = new List<Department>
                    {
                        new Department
                        {
                            DeptName = "Research & Development",
                            Employees = new List<Employee>
                            {
                                new Employee { FullName = "Alice Johnson", Age = 34 },
                                new Employee { FullName = "Bob Smith", Age = 29 }
                            }
                        },
                        new Department
                        {
                            DeptName = "Sales",
                            Employees = new List<Employee>
                            {
                                new Employee { FullName = "Carol White", Age = 41 },
                                new Employee { FullName = "David Brown", Age = 38 }
                            }
                        }
                    }
                }
            };

            return model;
        }
    }
}
