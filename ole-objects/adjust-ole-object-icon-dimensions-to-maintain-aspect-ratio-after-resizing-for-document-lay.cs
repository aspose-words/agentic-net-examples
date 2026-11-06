using System;

public class Program
{
    public static void Main()
    {
        // Original OLE object icon dimensions (in points)
        double originalWidth = 200.0;
        double originalHeight = 120.0;

        // Desired new width after layout resizing
        double newWidth = 150.0;

        // Calculate new height to maintain aspect ratio
        double aspectRatio = originalHeight / originalWidth;
        double newHeight = newWidth * aspectRatio;

        Console.WriteLine("OLE Object Icon Resize Example");
        Console.WriteLine("--------------------------------");
        Console.WriteLine($"Original Size : {originalWidth} x {originalHeight}");
        Console.WriteLine($"Requested Width: {newWidth}");
        Console.WriteLine($"Calculated Height (maintaining aspect ratio): {newHeight:F2}");
        Console.WriteLine($"Resized Size : {newWidth} x {newHeight:F2}");
    }
}
