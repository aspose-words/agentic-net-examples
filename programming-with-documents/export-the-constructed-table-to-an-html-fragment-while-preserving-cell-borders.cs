using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table and keep a reference to it.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndTable(); // Ends the table and moves the cursor after it.

        // Apply a simple black border to all sides of the table.
        table.SetBorders(LineStyle.Single, 1.0, Color.Black);

        // Export only the table to an HTML fragment.
        // The Table.ToString method renders the node using the specified format.
        string htmlFragment = table.ToString(SaveFormat.Html);

        // Define output path for the HTML fragment.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "TableFragment.html");

        // Save the fragment to a file.
        File.WriteAllText(outputPath, htmlFragment);

        // Optionally, write the fragment to the console.
        Console.WriteLine("HTML fragment saved to: " + outputPath);
        Console.WriteLine("Fragment content:");
        Console.WriteLine(htmlFragment);
    }
}
