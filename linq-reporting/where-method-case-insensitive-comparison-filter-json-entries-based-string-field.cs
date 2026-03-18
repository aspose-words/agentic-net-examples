using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;

// Define a model that matches the JSON structure.
public class Entry
{
    public string Name { get; set; }
    public int Id { get; set; }
    // Add other properties as needed.
}

public class Program
{
    public static void Main()
    {
        // Sample JSON array.
        string json = @"
        [
            { ""Name"": ""Alice"", ""Id"": 1 },
            { ""Name"": ""bob"",   ""Id"": 2 },
            { ""Name"": ""CHARLIE"", ""Id"": 3 },
            { ""Name"": ""alice"", ""Id"": 4 }
        ]";

        // Deserialize the JSON into a list of Entry objects.
        List<Entry> entries = JsonSerializer.Deserialize<List<Entry>>(json);

        // The value we want to filter by (case‑insensitive).
        string filterValue = "alice";

        // Use LINQ Where with StringComparison.OrdinalIgnoreCase for case‑insensitive matching.
        IEnumerable<Entry> filtered = entries
            .Where(e => string.Equals(e.Name, filterValue, StringComparison.OrdinalIgnoreCase));

        // Output the filtered results.
        foreach (var entry in filtered)
        {
            Console.WriteLine($"Id: {entry.Id}, Name: {entry.Name}");
        }
    }
}
