using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Root data model for the report.
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
        public string Name { get; set; } = string.Empty;
        public List<Project> Projects { get; set; } = new();
    }

    public class Project
    {
        public string Name { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create the template document programmatically.
            var templatePath = "Template.docx";
            var builder = new DocumentBuilder();
            // Department band.
            builder.Writeln("<<foreach [dept in Departments]>>");
            builder.Writeln("Department: <<[dept.Name]>>");
            builder.Writeln("");
            // Employee band inside department.
            builder.Writeln("<<foreach [emp in dept.Employees]>>");
            builder.Writeln("  Employee: <<[emp.Name]>>");
            builder.Writeln("");
            // Project band inside employee.
            builder.Writeln("  <<foreach [proj in emp.Projects]>>");
            builder.Writeln("    - <<[proj.Name]>>");
            builder.Writeln("  <</foreach>>");
            builder.Writeln("<</foreach>>");
            builder.Writeln("<</foreach>>");
            // Save the template.
            builder.Document.Save(templatePath);

            // 2. Load the template for reporting.
            var doc = new Document(templatePath);

            // 3. Prepare sample data.
            var model = new ReportModel
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
                                Name = "Alice Johnson",
                                Projects = new List<Project>
                                {
                                    new Project { Name = "AI Platform" },
                                    new Project { Name = "Data Lake" }
                                }
                            },
                            new Employee
                            {
                                Name = "Bob Smith",
                                Projects = new List<Project>
                                {
                                    new Project { Name = "IoT Framework" }
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
                                Name = "Carol White",
                                Projects = new List<Project>
                                {
                                    new Project { Name = "Social Media Campaign" },
                                    new Project { Name = "Brand Redesign" }
                                }
                            }
                        }
                    }
                }
            };

            // 4. Build the report using the LINQ Reporting engine.
            var engine = new ReportingEngine();
            // No special options are required for this simple scenario.
            engine.BuildReport(doc, model, "model");

            // 5. Save the generated report.
            doc.Save("Report.docx");
        }
    }
}
