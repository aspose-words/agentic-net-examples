using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a heading for the summary.
        builder.Writeln("Summary:");

        // Insert the "Total Sales" label.
        builder.Write("Total Sales: ");

        // Insert an expression field that will calculate the total sales.
        // The field code uses the Word calculation syntax.
        // The placeholder result "0" will be replaced when the field is updated.
        builder.InsertField("= SUM(Orders) \\* Arabic", "0");
        builder.Writeln();

        // Insert the "Average Order Value" label.
        builder.Write("Average Order Value: ");

        // Insert an expression field that will calculate the average order value.
        builder.InsertField("= AVERAGE(Orders) \\* Arabic", "0");
        builder.Writeln();

        // Update all fields so the calculations are performed (optional at this stage).
        doc.UpdateFields();

        // Save the document to disk.
        doc.Save("Summary.docx");
    }
}
