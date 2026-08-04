using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    public class Program
    {
        public static void Main()
        {
            // Enable reflection optimization for the reporting engine.
            ReportingEngine.UseReflectionOptimization = true;

            // Prepare hierarchical sample data.
            var data = new ReportData
            {
                Companies = new List<Company>
                {
                    new Company
                    {
                        Name = "Contoso Ltd.",
                        Departments = new List<Department>
                        {
                            new Department
                            {
                                Name = "Sales",
                                Employees = new List<Employee>
                                {
                                    new Employee { FirstName = "John", LastName = "Doe", Salary = 60000 },
                                    new Employee { FirstName = "Jane", LastName = "Smith", Salary = 65000 }
                                }
                            },
                            new Department
                            {
                                Name = "Development",
                                Employees = new List<Employee>
                                {
                                    new Employee { FirstName = "Alice", LastName = "Brown", Salary = 90000 },
                                    new Employee { FirstName = "Bob", LastName = "Johnson", Salary = 95000 }
                                }
                            }
                        }
                    },
                    new Company
                    {
                        Name = "Fabrikam Inc.",
                        Departments = new List<Department>
                        {
                            new Department
                            {
                                Name = "HR",
                                Employees = new List<Employee>
                                {
                                    new Employee { FirstName = "Emily", LastName = "Davis", Salary = 50000 }
                                }
                            }
                        }
                    }
                }
            };

            // Create a template document with LINQ Reporting tags.
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("<<foreach [c in Companies]>>");
            builder.Writeln("Company: <<[c.Name]>>");
            builder.Writeln("<<foreach [d in c.Departments]>>");
            builder.Writeln("  Department: <<[d.Name]>>");
            builder.Writeln("  <<foreach [e in d.Employees]>>");
            builder.Writeln("    - <<[e.FirstName]>> <<[e.LastName]>> : <<[e.Salary]>>");
            builder.Writeln("  <</foreach>>");
            builder.Writeln("<</foreach>>");
            builder.Writeln("<</foreach>>");

            // Initialize the reporting engine and register external types.
            var engine = new ReportingEngine();
            engine.KnownTypes.Add(typeof(Company));
            engine.KnownTypes.Add(typeof(Department));
            engine.KnownTypes.Add(typeof(Employee));

            // Build the report using the data root named "data".
            engine.BuildReport(doc, data, "data");

            // Save the generated report.
            doc.Save("ReportOutput.docx");
        }
    }

    // Root wrapper class.
    public class ReportData
    {
        public List<Company> Companies { get; set; } = new();
    }

    public class Company
    {
        public string Name { get; set; } = "";
        public List<Department> Departments { get; set; } = new();
    }

    public class Department
    {
        public string Name { get; set; } = "";
        public List<Employee> Employees { get; set; } = new();
    }

    public class Employee
    {
        public string FirstName { get; set; } = "";
        public string LastName { get; set; } = "";
        public double Salary { get; set; }
    }
}
