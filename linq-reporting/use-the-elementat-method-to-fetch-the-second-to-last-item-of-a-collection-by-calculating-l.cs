using System;
using System.Collections.Generic;
using System.Linq;

public class Program
{
    public static void Main()
    {
        // Sample collection of integers.
        List<int> numbers = new List<int> { 10, 20, 30, 40, 50, 60 };

        // Calculate the index of the second-to-last element.
        int secondToLastIndex = numbers.Count - 2;

        // Use ElementAt to fetch the element at that index.
        int secondToLast = numbers.ElementAt(secondToLastIndex);

        // Output the result.
        Console.WriteLine($"Second-to-last element: {secondToLast}");
    }
}
