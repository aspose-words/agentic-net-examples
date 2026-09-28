using System;

public class Program
{
    public static void Main()
    {
        // Simulated original image dimensions (e.g., pixels)
        int originalWidth = 800;
        int originalHeight = 600;

        // Create merging arguments and set the desired scale
        var args = new ImageFieldMergingArgs
        {
            FieldName = "CompanyLogo",
            ImageData = new byte[0], // Placeholder for image bytes
            ImageScale = 0.5 // Scale to 50%
        };

        // Adjust image size based on the ImageScale property
        int scaledWidth = (int)(originalWidth * args.ImageScale);
        int scaledHeight = (int)(originalHeight * args.ImageScale);

        Console.WriteLine($"Merging field '{args.FieldName}' with image scaled to {args.ImageScale * 100}%.");
        Console.WriteLine($"Original size: {originalWidth}x{originalHeight}");
        Console.WriteLine($"Scaled size: {scaledWidth}x{scaledHeight}");
    }
}

public class ImageFieldMergingArgs
{
    public string FieldName { get; set; }
    public byte[] ImageData { get; set; }
    public double ImageScale { get; set; } = 1.0;
}
