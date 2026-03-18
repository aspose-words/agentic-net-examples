using System;
using System.Drawing;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Example numeric data.
        int[] numbers = { 12, 45, 67, 82, 95 };

        foreach (int number in numbers)
        {
            // Determine the text color based on thresholds.
            Color color;
            if (number < 30)
                color = Color.Red;          // Low values -> red.
            else if (number < 70)
                color = Color.Orange;       // Medium values -> orange.
            else
                color = Color.Green;        // High values -> green.

            // Apply the selected color.
            builder.Font.Color = color;

            // Write the number with the appropriate color.
            builder.Writeln($"Number: {number}");
        }

        // Save the resulting document.
        doc.Save("DynamicColor.docx");
    }
}
