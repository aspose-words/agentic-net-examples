using System;
using System.Collections.Generic;
using System.Linq;

public class Employee
{
    // Seniority level of the employee (higher means more senior).
    public int Seniority { get; set; }

    // Employee's name.
    public string Name { get; set; }

    public Employee(int seniority, string name)
    {
        Seniority = seniority;
        Name = name;
    }
}

public class Program
{
    public static void Main()
    {
        // Example collection of employees.
        List<Employee> employees = new List<Employee>
        {
            new Employee(5, "Alice"),
            new Employee(3, "Bob"),
            new Employee(5, "Charlie"),
            new Employee(2, "David")
        };

        // Chain OrderByDescending (seniority) and ThenBy (name) to sort.
        List<Employee> sortedEmployees = employees
            .OrderByDescending(e => e.Seniority) // Highest seniority first.
            .ThenBy(e => e.Name)                 // Within same seniority, sort alphabetically.
            .ToList();

        // Display the sorted list.
        foreach (Employee emp in sortedEmployees)
        {
            Console.WriteLine($"{emp.Seniority} - {emp.Name}");
        }
    }
}
