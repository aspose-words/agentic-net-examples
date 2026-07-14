using System;
using System.Collections.Generic;

public class Program
{
    public static void Main()
    {
        var fields = new List<string> { "Logo", "Signature", "Photo", "Unknown" };
        foreach (var field in fields)
        {
            string imagePath = ImageFieldMerging(field);
            Console.WriteLine($"Field: {field}, Image: {imagePath}");
        }
    }

    private static string ImageFieldMerging(string fieldName)
    {
        // Conditional logic to select image based on field name
        if (string.Equals(fieldName, "Logo", StringComparison.OrdinalIgnoreCase))
        {
            return "Images/CompanyLogo.png";
        }
        else if (string.Equals(fieldName, "Signature", StringComparison.OrdinalIgnoreCase))
        {
            return "Images/Signature.png";
        }
        else if (string.Equals(fieldName, "Photo", StringComparison.OrdinalIgnoreCase))
        {
            return "Images/EmployeePhoto.jpg";
        }
        else
        {
            return "Images/Default.png";
        }
    }
}
