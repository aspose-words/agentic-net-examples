using System;
using System.Linq;
using System.Collections.Generic;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load an existing Word document.
        Document doc = new Document("Input.docx");

        // Retrieve all Run nodes whose font size is greater than 12 points.
        // The GetChildNodes method returns a live collection of nodes.
        // Cast to Run and filter with a lambda expression.
        List<Run> runsWithLargeFont = doc.GetChildNodes(NodeType.Run, true)
                                         .Cast<Run>()
                                         .Where(r => r.Font.Size > 12)
                                         .ToList();

        // Example usage: write the text of each matching run to the console.
        foreach (Run run in runsWithLargeFont)
        {
            Console.WriteLine(run.Text);
        }

        // Save the (potentially unchanged) document to a new file.
        doc.Save("Output.docx");
    }
}
