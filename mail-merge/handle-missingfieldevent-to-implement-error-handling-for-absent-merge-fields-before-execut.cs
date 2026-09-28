using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

public class MailMergeEngine
{
    public delegate void MissingFieldEventHandler(string fieldName);
    public event MissingFieldEventHandler MissingField;

    public string Process(string template, Dictionary<string, string> data)
    {
        if (template == null) throw new ArgumentNullException(nameof(template));
        if (data == null) throw new ArgumentNullException(nameof(data));

        // Pattern matches {{FieldName}}
        var pattern = new Regex(@"{{\s*(\w+)\s*}}", RegexOptions.Compiled);
        var result = pattern.Replace(template, match =>
        {
            var fieldName = match.Groups[1].Value;
            if (data.TryGetValue(fieldName, out var value))
            {
                return value;
            }
            else
            {
                // Raise event for missing field
                MissingField?.Invoke(fieldName);
                // After event, try again
                if (data.TryGetValue(fieldName, out var newValue))
                {
                    return newValue;
                }
                // If still missing, keep placeholder unchanged
                return match.Value;
            }
        });

        return result;
    }
}

public class Program
{
    public static void Main()
    {
        var template = "Dear {{FirstName}} {{LastName}},\nYour order {{OrderId}} is shipped.";
        var data = new Dictionary<string, string>
        {
            { "FirstName", "John" },
            { "OrderId", "12345" }
            // Note: LastName is intentionally missing
        };

        var engine = new MailMergeEngine();

        // Subscribe to MissingField event to provide a default value
        engine.MissingField += fieldName =>
        {
            // Provide a default placeholder for any missing field
            data[fieldName] = "[Missing]";
        };

        var result = engine.Process(template, data);
        Console.WriteLine(result);
    }
}
