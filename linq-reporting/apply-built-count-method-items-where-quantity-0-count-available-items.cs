using System;
using System.Collections.Generic;
using System.Linq;

class Item
{
    public int Quantity { get; set; }
}

class Program
{
    static void Main()
    {
        // Example collection of items.
        List<Item> items = new List<Item>
        {
            new Item { Quantity = 5 },
            new Item { Quantity = 0 },
            new Item { Quantity = 3 },
            new Item { Quantity = -2 }
        };

        // Count items where Quantity > 0.
        int availableCount = items.Count(item => item.Quantity > 0);

        Console.WriteLine($"Available items: {availableCount}");
    }
}
