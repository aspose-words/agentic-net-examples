using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

namespace LinqReportingExample
{
    // Employee data model
    public class Employee
    {
        public int Seniority { get; set; }
        public string Name { get; set; } = string.Empty;
    }

    // Root model containing the collection of employees
    public class ReportModel
    {
        public List<Employee> Employees { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for Aspose.Words)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample data
            var model = new ReportModel
            {
                Employees = new List<Employee>
                {
                    new Employee { Seniority = 5, Name = "Alice" },
                    new Employee { Seniority = 3, Name = "Bob" },
                    new Employee { Seniority = 5, Name = "Charlie" },
                    new Employee { Seniority = 2, Name = "Diana" }
                }
            };

            // Create a template document programmatically
            var templatePath = "Template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Insert LINQ Reporting tags with sorting: OrderByDescending then ThenBy
            builder.Writeln("<<foreach [e in Employees.OrderByDescending(e => e.Seniority).ThenBy(e => e.Name)]>>");
            builder.Writeln("Name: <<[e.Name]>>, Seniority: <<[e.Seniority]>>");
            builder.Writeln("<</foreach>>");

            // Save the template
            doc.Save(templatePath);

            // Load the template for report generation
            var reportDoc = new Document(templatePath);

            // Build the report
            var engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report
            var outputPath = "Report.docx";
            reportDoc.Save(outputPath);
        }
    }
}
