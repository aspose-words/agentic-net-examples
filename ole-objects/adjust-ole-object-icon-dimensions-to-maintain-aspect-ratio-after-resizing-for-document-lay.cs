using System;

public class Program
{
    // Adjusts the height to maintain aspect ratio when the width is changed.
    private static double GetAdjustedHeight(double originalWidth, double originalHeight, double newWidth)
    {
        if (originalWidth == 0) throw new ArgumentException("Original width cannot be zero.");
        double aspectRatio = originalHeight / originalWidth;
        return newWidth * aspectRatio;
    }

    // Adjusts the width to maintain aspect ratio when the height is changed.
    private static double GetAdjustedWidth(double originalWidth, double originalHeight, double newHeight)
    {
        if (originalHeight == 0) throw new ArgumentException("Original height cannot be zero.");
        double aspectRatio = originalWidth / originalHeight;
        return newHeight * aspectRatio;
    }

    public static void Main()
    {
        // Example original OLE icon dimensions (in points)
        double originalWidth = 120.0;
        double originalHeight = 80.0;

        // Desired new width after resizing the document layout
        double newWidth = 200.0;

        // Compute the new height to keep the aspect ratio
        double adjustedHeight = GetAdjustedHeight(originalWidth, originalHeight, newWidth);

        // Output the results
        Console.WriteLine("Original dimensions: {0} x {1}", originalWidth, originalHeight);
        Console.WriteLine("Requested new width: {0}", newWidth);
        Console.WriteLine("Adjusted height to maintain aspect ratio: {0:F2}", adjustedHeight);

        // Additional example: resizing based on a new height
        double newHeight = 150.0;
        double adjustedWidth = GetAdjustedWidth(originalWidth, originalHeight, newHeight);
        Console.WriteLine();
        Console.WriteLine("Requested new height: {0}", newHeight);
        Console.WriteLine("Adjusted width to maintain aspect ratio: {0:F2}", adjustedWidth);
    }
}
