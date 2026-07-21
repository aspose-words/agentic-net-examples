using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add sample lists to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First list – numbered.
        builder.Writeln("Numbered list:");
        builder.ListFormat.ApplyNumberDefault();
        for (int i = 1; i <= 3; i++)
        {
            builder.Writeln($"Item {i}");
        }
        builder.ListFormat.RemoveNumbers();

        // Second list – bulleted.
        builder.Writeln("\nBulleted list:");
        builder.ListFormat.ApplyBulletDefault();
        for (int i = 1; i <= 3; i++)
        {
            builder.Writeln($"Bullet {i}");
        }
        builder.ListFormat.RemoveNumbers();

        // Iterate over all lists in the document and apply a uniform style to each level.
        foreach (List list in doc.Lists)
        {
            foreach (ListLevel level in list.ListLevels)
            {
                // Set the desired uniform formatting.
                level.Font.Name = "Arial";
                level.Font.Color = Color.Green;
                level.Font.Bold = true;
                // Optionally, you can set other properties such as NumberStyle, NumberFormat, etc.
            }
        }

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "UniformListStyle.docx");
        doc.Save(outputPath);
    }
}
