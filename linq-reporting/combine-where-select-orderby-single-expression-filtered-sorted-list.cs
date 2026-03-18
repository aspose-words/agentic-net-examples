using System;
using System.Collections.Generic;
using System.Linq;

namespace LinqCombinedExample
{
    // Simple data entity.
    public class Customer
    {
        public string FirstName { get; set; }
        public string LastName  { get; set; }
        public int    Age       { get; set; }

        public Customer(string firstName, string lastName, int age)
        {
            FirstName = firstName;
            LastName  = lastName;
            Age       = age;
        }
    }

    class Program
    {
        static void Main()
        {
            // Sample data source.
            var customers = new List<Customer>
            {
                new Customer("John",   "Doe",      25),
                new Customer("Jane",   "Smith",    30),
                new Customer("Alice",  "Brown",    17),
                new Customer("Bob",    "Johnson",  45),
                new Customer("Carol",  "Davis",    22)
            };

            // One‑liner that filters (Age > 20), sorts by LastName ascending,
            // and projects to an anonymous type containing only the names.
            var filteredSorted = customers
                .Where(c => c.Age > 20)                     // filter
                .OrderBy(c => c.LastName)                   // sort
                .Select(c => new { c.FirstName, c.LastName }) // project
                .ToList();                                  // materialize

            // Output the result.
            foreach (var entry in filteredSorted)
                Console.WriteLine($"{entry.FirstName} {entry.LastName}");
        }
    }
}
