using System;
using System.Collections.Generic;
using System.Linq;

public class Program
{
    public static void Main()
    {
        // Sample collection.
        List<string> fruits = new List<string>
        {
            "Apple",
            "Banana",
            "Cherry",
            "Date",
            "Elderberry"
        };

        // Calculate the index of the second‑to‑last element.
        int secondToLastIndex = fruits.Count - 2;

        // Use ElementAt to fetch the element at that index.
        string secondToLastFruit = fruits.ElementAt(secondToLastIndex);

        // Output the result.
        Console.WriteLine($"Second‑to‑last fruit: {secondToLastFruit}");
    }
}
