using System;

public class Program
{
    public static void Main()
    {
        // Original OLE icon dimensions (e.g., pixels)
        double originalWidth = 200;
        double originalHeight = 100;

        // Resize based on a new width while preserving aspect ratio
        double targetWidth = 150;
        var resizedByWidth = OLEIconResizer.ResizeByWidth(originalWidth, originalHeight, targetWidth);
        Console.WriteLine($"Original size: {originalWidth} x {originalHeight}");
        Console.WriteLine($"Resized to width {targetWidth}: {resizedByWidth.Width:F2} x {resizedByWidth.Height:F2}");

        // Resize based on a new height while preserving aspect ratio
        double targetHeight = 80;
        var resizedByHeight = OLEIconResizer.ResizeByHeight(originalWidth, originalHeight, targetHeight);
        Console.WriteLine($"Resized to height {targetHeight}: {resizedByHeight.Width:F2} x {resizedByHeight.Height:F2}");
    }
}

public static class OLEIconResizer
{
    public struct Size
    {
        public double Width;
        public double Height;

        public Size(double width, double height)
        {
            Width = width;
            Height = height;
        }
    }

    // Adjust dimensions based on a new width, maintaining aspect ratio
    public static Size ResizeByWidth(double originalWidth, double originalHeight, double newWidth)
    {
        if (originalWidth == 0) return new Size(0, 0);
        double scale = newWidth / originalWidth;
        double newHeight = originalHeight * scale;
        return new Size(newWidth, newHeight);
    }

    // Adjust dimensions based on a new height, maintaining aspect ratio
    public static Size ResizeByHeight(double originalWidth, double originalHeight, double newHeight)
    {
        if (originalHeight == 0) return new Size(0, 0);
        double scale = newHeight / originalHeight;
        double newWidth = originalWidth * scale;
        return new Size(newWidth, newHeight);
    }
}
