using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Initial paragraph.");

        // Start tracking revisions.
        doc.StartTrackRevisions("DemoUser", DateTime.Now);

        // Insert a 2x2 table while tracking is enabled.
        builder.StartTable();

        for (int row = 1; row <= 2; row++)
        {
            for (int col = 1; col <= 2; col++)
            {
                builder.InsertCell();
                builder.Write($"Cell {row},{col}");
            }

            builder.EndRow();
        }

        builder.EndTable();

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Accept the table insertion revision(s).
        foreach (Revision rev in doc.Revisions)
        {
            // The insertion of the table itself is represented by a revision whose
            // ParentNode is a Table node.
            if (rev.RevisionType == RevisionType.Insertion && rev.ParentNode is Table)
            {
                rev.Accept();
            }
        }

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
