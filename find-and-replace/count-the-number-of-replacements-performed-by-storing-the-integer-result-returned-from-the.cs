using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and add sample text containing placeholders.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Dear _Customer_,");
        builder.Writeln("Your order number _OrderNumber_ has been shipped.");
        builder.Writeln("Thank you for shopping with us.");

        // Replace the first placeholder and store the number of replacements made.
        int customerReplacements = doc.Range.Replace("_Customer_", "Alice");

        // Replace the second placeholder and store the number of replacements made.
        int orderReplacements = doc.Range.Replace("_OrderNumber_", "12345");

        // Sum the replacement counts.
        int totalReplacements = customerReplacements + orderReplacements;

        // Ensure that at least one replacement occurred.
        if (totalReplacements == 0)
            throw new InvalidOperationException("No replacements were performed.");

        // Save the modified document to a local file.
        const string outputPath = "ModifiedDocument.docx";
        doc.Save(outputPath);

        // Output the total number of replacements to the console.
        Console.WriteLine($"Total replacements performed: {totalReplacements}");
    }
}
