using System;
using System.Collections.Generic;
using System.Linq;

namespace AsposeWordsExample
{
    // Simple data model representing an employee.
    public class Employee
    {
        public string Name { get; set; }
        public int Age { get; set; }

        public Employee(string name, int age)
        {
            Name = name;
            Age = age;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare a sample collection of employees.
            List<Employee> employees = new List<Employee>
            {
                new Employee("Alice", 28),
                new Employee("Bob", 35),
                new Employee("Charlie", 42)
            };

            // Use LINQ's FirstOrDefault to retrieve the first employee older than 30.
            // If no employee satisfies the predicate, the result will be null.
            Employee firstOlderThan30 = employees.FirstOrDefault(e => e.Age > 30);

            if (firstOlderThan30 != null)
            {
                Console.WriteLine($"First employee older than 30: {firstOlderThan30.Name}, Age {firstOlderThan30.Age}");
            }
            else
            {
                Console.WriteLine("No employee older than 30 was found.");
            }

            // Example of retrieving the very first employee regardless of condition.
            Employee firstEmployee = employees.FirstOrDefault();

            if (firstEmployee != null)
            {
                Console.WriteLine($"First employee in the collection: {firstEmployee.Name}, Age {firstEmployee.Age}");
            }
        }
    }
}
