using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsWhereExample
{
    // Simple data entity.
    public class Employee
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    // Wrapper that holds the collection and the dynamic filter value.
    public class ReportData
    {
        public int MinAge { get; set; }
        public List<Employee> Employees { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a template document in memory.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Header.
            builder.Writeln("Employees with Age > MinAge:");
            // ReportingEngine template syntax using the LINQ 'where' extension method.
            // The placeholder <<foreach [Employees where Age > MinAge]>> iterates only over
            // employees that satisfy the runtime condition (Age > MinAge).
            builder.Writeln("<<foreach [Employees where Age > MinAge]>>");
            builder.Writeln("- <<[Name]>> (Age: <<[Age]>>)");
            builder.Writeln("<<end>>");

            // 2. Prepare the data source.
            var data = new ReportData
            {
                MinAge = 30, // Dynamic runtime criteria.
                Employees = new List<Employee>
                {
                    new Employee { Name = "John Doe", Age = 28 },
                    new Employee { Name = "Jane Smith", Age = 35 },
                    new Employee { Name = "Bob Johnson", Age = 42 },
                    new Employee { Name = "Alice Brown", Age = 25 }
                }
            };

            // 3. Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // The data source object is passed directly; the template can reference its members.
            engine.BuildReport(template, data);

            // 4. Save the resulting document.
            template.Save("EmployeesReport.docx");
        }
    }
}
