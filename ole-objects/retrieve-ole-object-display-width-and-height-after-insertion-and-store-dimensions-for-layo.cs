using System;
using System.Collections.Generic;

public class OLEObject
{
    public string Name { get; }
    public double Width { get; private set; }   // Width in points
    public double Height { get; private set; }  // Height in points

    // Simulate insertion which determines display size
    public OLEObject(string name, double intrinsicWidth, double intrinsicHeight)
    {
        Name = name;
        // In a real scenario, the display size might be calculated based on the object's
        // intrinsic size and the host application's scaling factors.
        // Here we simply assign the intrinsic dimensions as the display size.
        Width = intrinsicWidth;
        Height = intrinsicHeight;
    }

    // Simulate a method that could adjust size after insertion (e.g., user resizing)
    public void Resize(double newWidth, double newHeight)
    {
        Width = newWidth;
        Height = newHeight;
    }
}

public class Program
{
    public static void Main()
    {
        // Simulate inserting several OLE objects into a document
        var insertedObjects = new List<OLEObject>
        {
            new OLEObject("Chart1", 300, 200),
            new OLEObject("Image1", 400, 250),
            new OLEObject("ExcelSheet1", 500, 350)
        };

        // After insertion, retrieve display dimensions
        var layoutDimensions = new Dictionary<string, (double Width, double Height)>();

        foreach (var ole in insertedObjects)
        {
            // In a real implementation, you would query the host (e.g., Word) for the actual
            // displayed size. Here we use the properties directly.
            layoutDimensions[ole.Name] = (ole.Width, ole.Height);
        }

        // Example layout calculation: total occupied width and height
        double totalWidth = 0;
        double maxHeight = 0;

        foreach (var kvp in layoutDimensions)
        {
            totalWidth += kvp.Value.Width;
            if (kvp.Value.Height > maxHeight)
                maxHeight = kvp.Value.Height;
        }

        // Output the collected dimensions and layout summary
        Console.WriteLine("OLE Object Dimensions:");
        foreach (var kvp in layoutDimensions)
        {
            Console.WriteLine($"- {kvp.Key}: Width = {kvp.Value.Width} pt, Height = {kvp.Value.Height} pt");
        }

        Console.WriteLine();
        Console.WriteLine($"Total width of all objects: {totalWidth} pt");
        Console.WriteLine($"Maximum height among objects: {maxHeight} pt");
    }
}
