using System;
using System.Collections.Generic;

public class ListDefinition
{
    public string Name { get; set; }
    public int Level { get; set; }

    public ListDefinition(string name, int level)
    {
        Name = name;
        Level = level;
    }

    public override string ToString()
    {
        return $"{Name} (Level {Level})";
    }
}

public class Document
{
    public List<ListDefinition> Lists { get; } = new List<ListDefinition>();

    public Document()
    {
        // Sample list definitions
        Lists.Add(new ListDefinition("BulletList", 1));
        Lists.Add(new ListDefinition("NumberedList", 2));
        Lists.Add(new ListDefinition("OutlineList", 3));
    }
}

public class Program
{
    public static void Main()
    {
        // Create a document with predefined lists
        Document doc = new Document();

        Console.WriteLine("Before modification:");
        foreach (var list in doc.Lists)
        {
            Console.WriteLine(list);
        }

        // Uniformly modify all list definitions (e.g., set Level to 1)
        foreach (var list in doc.Lists)
        {
            list.Level = 1;
        }

        Console.WriteLine("\nAfter modification:");
        foreach (var list in doc.Lists)
        {
            Console.WriteLine(list);
        }
    }
}
