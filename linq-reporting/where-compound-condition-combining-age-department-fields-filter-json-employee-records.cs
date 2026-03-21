using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;

namespace AsposeWordsJsonFilter
{
    // Simple POCO that matches the structure of each employee record in the JSON data.
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
            // Sample JSON data representing a collection of employee records.
            string jsonContent = @"
            [
                { ""Name"": ""Alice"",   ""Age"": 35, ""Department"": ""Sales"" },
                { ""Name"": ""Bob"",     ""Age"": 28, ""Department"": ""HR"" },
                { ""Name"": ""Charlie"", ""Age"": 40, ""Department"": ""Sales"" },
                { ""Name"": ""Diana"",   ""Age"": 32, ""Department"": ""IT"" }
            ]";

            // Deserialize the JSON array into a list of Employee objects.
            List<Employee> allEmployees = JsonSerializer.Deserialize<List<Employee>>(jsonContent);

            // Apply a compound LINQ Where clause: Age > 30 AND Department == "Sales".
            IEnumerable<Employee> filteredEmployees = allEmployees
                .Where(e => e.Age > 30 && e.Department == "Sales");

            // Output the filtered results to the console.
            Console.WriteLine("Filtered Employees (Age > 30 && Department == \"Sales\"):");
            foreach (Employee emp in filteredEmployees)
            {
                Console.WriteLine($"- {emp.Name}, Age: {emp.Age}, Department: {emp.Department}");
            }
        }
    }
}
