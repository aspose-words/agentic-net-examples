using System;
using System.Collections.Generic;
using System.Linq;

class Person
{
    public string First { get; }
    public string Middle { get; }
    public string Last { get; }

    public Person(string first, string middle, string last)
    {
        First = first;
        Middle = middle;
        Last = last;
    }
}

class RetrieveFourthPerson
{
    static void Main()
    {
        var persons = new List<Person>
        {
            new Person("Alice", "B.", "Carroll"),
            new Person("Bob", "C.", "Davis"),
            new Person("Carol", "D.", "Evans"),
            new Person("David", "E.", "Foster")
        };

        // Retrieve the fourth person (zero‑based index 3) using LINQ's ElementAt.
        Person fourthPerson = persons.ElementAt(3);

        // Display the person's full name.
        Console.WriteLine($"Fourth author: {fourthPerson.First} {fourthPerson.Middle} {fourthPerson.Last}");
    }
}
