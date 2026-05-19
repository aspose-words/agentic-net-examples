using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Data model classes
    public class ReportModel
    {
        public List<Department> Departments { get; set; } = new();
    }

    public class Department
    {
        public string Name { get; set; } = string.Empty;
        public List<Employee> Employees { get; set; } = new();
    }

    public class Employee
    {
        public string FullName { get; set; } = string.Empty;
        public List<Project> Projects { get; set; } = new();
    }

    public class Project
    {
        public string Title { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a template document programmatically.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Title
            builder.Writeln("Company Project Report");
            builder.Writeln();

            // Outer loop: Departments
            builder.Writeln("<<foreach [dept in Departments]>>");
            builder.Writeln("Department: <<[dept.Name]>>");
            builder.Writeln();

            // Middle loop: Employees within a department
            builder.Writeln("<<foreach [emp in dept.Employees]>>");
            builder.Writeln("  Employee: <<[emp.FullName]>>");
            builder.Writeln();

            // Inner loop: Projects for each employee
            builder.Writeln("  <<foreach [proj in emp.Projects]>>");
            builder.Writeln("    Project: <<[proj.Title]>> - <<[proj.Description]>>");
            builder.Writeln("  <</foreach>>");
            builder.Writeln();
            builder.Writeln("<</foreach>>");
            builder.Writeln();
            builder.Writeln("<</foreach>>");

            // 2. Prepare sample data.
            ReportModel model = new ReportModel
            {
                Departments = new List<Department>
                {
                    new Department
                    {
                        Name = "Research & Development",
                        Employees = new List<Employee>
                        {
                            new Employee
                            {
                                FullName = "Alice Johnson",
                                Projects = new List<Project>
                                {
                                    new Project { Title = "AI Platform", Description = "Develop core AI services." },
                                    new Project { Title = "Data Pipeline", Description = "Build scalable data ingestion." }
                                }
                            },
                            new Employee
                            {
                                FullName = "Bob Smith",
                                Projects = new List<Project>
                                {
                                    new Project { Title = "Quantum Research", Description = "Explore quantum algorithms." }
                                }
                            }
                        }
                    },
                    new Department
                    {
                        Name = "Marketing",
                        Employees = new List<Employee>
                        {
                            new Employee
                            {
                                FullName = "Carol White",
                                Projects = new List<Project>
                                {
                                    new Project { Title = "Brand Refresh", Description = "Update visual identity." },
                                    new Project { Title = "Social Campaign", Description = "Launch Q3 social media ads." }
                                }
                            }
                        }
                    }
                }
            };

            // 3. Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // 4. Save the generated report.
            const string outputPath = "NestedReport.docx";
            doc.Save(outputPath);
        }
    }
}
