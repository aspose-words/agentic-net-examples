using System;
using System.Collections.Generic;
using System.Linq;

public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Sample data
        var people = new List<Person>
        {
            new() { Name = "Alice",   Age = 30 },
            new() { Name = "Bob",     Age = 25 },
            new() { Name = "Charlie", Age = 35 },
            new() { Name = "David",   Age = 28 }
        };

        // Combine Where, Select, and OrderBy in a single LINQ expression
        var filteredSorted = people
            .Where(p => p.Age >= 28)                 // Filter: age 28 or older
            .Select(p => new { p.Name, p.Age })      // Project to anonymous type
            .OrderBy(p => p.Name);                   // Sort by name

        // Output the results
        foreach (var person in filteredSorted)
        {
            Console.WriteLine($"{person.Name} - {person.Age}");
        }
    }
}
