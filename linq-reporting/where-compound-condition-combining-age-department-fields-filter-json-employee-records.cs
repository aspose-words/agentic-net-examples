using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsJsonFilter
{
    // Simple POCO that matches the structure of each employee record in the JSON file.
    public class Employee
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public string Department { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Paths to the template document, the source JSON file and the output document.
            string templatePath = @"C:\Docs\EmployeeTemplate.docx";
            string jsonPath = @"C:\Data\employees.json";
            string outputPath = @"C:\Docs\FilteredEmployees.docx";

            // Load the JSON file content.
            string jsonContent = File.ReadAllText(jsonPath);

            // Deserialize the JSON array into a list of Employee objects.
            List<Employee> allEmployees = JsonSerializer.Deserialize<List<Employee>>(jsonContent);

            // Apply a compound LINQ Where clause: Age > 30 AND Department == "Sales".
            IEnumerable<Employee> filteredEmployees = allEmployees
                .Where(e => e.Age > 30 && e.Department == "Sales");

            // Convert the filtered collection into a DataTable for mail merge.
            DataTable employeeTable = new DataTable("Employees");
            employeeTable.Columns.Add("Name", typeof(string));
            employeeTable.Columns.Add("Age", typeof(int));
            employeeTable.Columns.Add("Department", typeof(string));

            foreach (Employee emp in filteredEmployees)
            {
                employeeTable.Rows.Add(emp.Name, emp.Age, emp.Department);
            }

            // Load the Word template document.
            Document doc = new Document(templatePath);

            // Perform mail merge using the filtered DataTable.
            doc.MailMerge.Execute(employeeTable);

            // Save the resulting document.
            doc.Save(outputPath);
        }
    }
}
