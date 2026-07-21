using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

public class MailMergeDemo
{
    public static void Main()
    {
        // Template with text and image placeholders
        string template = @"
Dear {{Name}},

Here is your logo: {{Image:Logo}}
Your signature: {{Image:Signature}}
Other image: {{Image:Other}}

Best regards,
Company";

        // Text fields
        var textFields = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "Name", "John Doe" }
        };

        // Perform text field merging
        string merged = MergeTextFields(template, textFields);

        // Perform image field merging with conditional logic
        merged = MergeImageFields(merged);

        // Output the final merged result
        Console.WriteLine(merged);
    }

    private static string MergeTextFields(string template, Dictionary<string, string> fields)
    {
        foreach (var kvp in fields)
        {
            string placeholder = $"{{{{{kvp.Key}}}}}";
            template = template.Replace(placeholder, kvp.Value);
        }
        return template;
    }

    private static string MergeImageFields(string template)
    {
        // Regex to find {{Image:FieldName}} placeholders
        var regex = new Regex(@"\{\{Image:(?<field>\w+)\}\}", RegexOptions.IgnoreCase);
        return regex.Replace(template, match =>
        {
            string fieldName = match.Groups["field"].Value;
            string imagePath = GetImagePath(fieldName);
            // In a real mail merge, you would embed the image.
            // For this demo, we just return the image file name.
            return imagePath;
        });
    }

    private static string GetImagePath(string fieldName)
    {
        // Conditional logic to select image based on field name
        switch (fieldName.ToLowerInvariant())
        {
            case "logo":
                return "logo.png";
            case "signature":
                return "signature.png";
            default:
                return "default.png";
        }
    }
}
