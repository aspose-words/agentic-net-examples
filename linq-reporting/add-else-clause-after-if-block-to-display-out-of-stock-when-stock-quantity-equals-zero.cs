using System;

public class Program
{
    public static void Main()
    {
        // Sample product with zero stock
        var product = new Product
        {
            Name = "Widget",
            Stock = 0
        };

        Console.WriteLine($"Product: {product.Name}");

        // Display stock status
        if (product.Stock > 0)
        {
            Console.WriteLine($"In stock: {product.Stock}");
        }
        else
        {
            // Stock is zero, show out‑of‑stock message
            Console.WriteLine("Out of stock");
        }
    }
}

// Simple data model for a product
public class Product
{
    // Product name (initialized to avoid nullable warnings)
    public string Name { get; set; } = string.Empty;

    // Quantity available in stock
    public int Stock { get; set; }
}
