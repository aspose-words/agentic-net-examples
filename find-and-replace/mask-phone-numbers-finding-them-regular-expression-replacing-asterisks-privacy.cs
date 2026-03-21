using System;
using System.Text.RegularExpressions;

class PhoneNumberMasking
{
    static void Main()
    {
        // Sample text containing various phone number formats.
        string input = @"Contact us at 123-456-7890 or (123) 456-7890. Also +1 123.456.7890 and 1234567890.";

        // Regular expression that matches common US phone number formats:
        // 123-456-7890, (123) 456-7890, 123 456 7890, 123.456.7890, 1234567890, +1 123-456-7890, etc.
        string phonePattern = @"\b(?:\+?1[\s\-\.]?)?(?:\(\d{3}\)|\d{3})[\s\-\.]?\d{3}[\s\-\.]?\d{4}\b";

        // Replace each found phone number with a string of asterisks of the same length.
        string masked = Regex.Replace(input, phonePattern, m => new string('*', m.Value.Length));

        Console.WriteLine("Original: " + input);
        Console.WriteLine("Masked:   " + masked);
    }
}
