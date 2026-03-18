using System;
using System.Drawing;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define a set of statuses we want to display.
        string[] statuses = { "Completed", "In Progress", "Failed" };

        // Iterate over each status, set the run color, and write a line.
        foreach (string status in statuses)
        {
            // Set the font color based on the status using a helper method.
            builder.Font.Color = GetColorForStatus(status);

            // Write the status line. Each call creates a new Run internally.
            builder.Writeln($"{status}: The operation has finished.");
        }

        // Save the resulting document to disk.
        doc.Save("RunTextColorByStatus.docx");
    }

    // Helper method that maps a status string to a specific Color.
    static Color GetColorForStatus(string status)
    {
        // Simple mapping; extend as needed for additional statuses.
        switch (status)
        {
            case "Completed":
                return Color.Green;      // Success -> green
            case "In Progress":
                return Color.Orange;     // Ongoing -> orange
            case "Failed":
                return Color.Red;        // Failure -> red
            default:
                return Color.Black;      // Fallback -> black
        }
    }
}
