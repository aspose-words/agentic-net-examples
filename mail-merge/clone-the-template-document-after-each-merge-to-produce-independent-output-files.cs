using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a template document in memory.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Dear <<Name>>,");
        builder.Writeln("Your order <<OrderNumber>> has been shipped.");
        builder.Writeln("Thank you for shopping with us.");

        // Data for mail merge.
        string[] names = { "John Doe", "Jane Smith", "Bob Johnson" };
        int[] orderNumbers = { 1001, 1002, 1003 };

        // Perform mail merge for each record, cloning the template each time.
        for (int i = 0; i < names.Length; i++)
        {
            // Clone the template to get an independent document.
            Document output = (Document)template.Clone(true);

            // Execute mail merge with a single record.
            output.MailMerge.Execute(
                new string[] { "Name", "OrderNumber" },
                new object[] { names[i], orderNumbers[i] });

            // Save the result to a separate file.
            string fileName = $"Output_{i + 1}.docx";
            output.Save(fileName);
        }
    }
}
