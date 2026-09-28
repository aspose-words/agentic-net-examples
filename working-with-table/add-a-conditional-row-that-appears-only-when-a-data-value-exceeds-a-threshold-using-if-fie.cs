using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a document variable that will be used in the IF field.
        doc.Variables.Add("Score", "15"); // Change this value to test the condition.

        // Start the table.
        builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Score");
        builder.EndRow();

        // Data row.
        builder.InsertCell();
        builder.Write("Test");
        builder.InsertCell();
        // Insert a DOCVARIABLE field to display the score.
        builder.InsertField("DOCVARIABLE Score");
        builder.EndRow();

        // Conditional row – appears only when Score > 10.
        builder.InsertCell();
        builder.Write("Note");
        builder.InsertCell();
        // IF field: if Score > 10, show the message; otherwise show nothing.
        builder.InsertField("IF { DOCVARIABLE Score } > 10 \"Exceeds threshold\" \"\"");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Update fields to evaluate the IF condition.
        doc.UpdateFields();

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ConditionalTable.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        // Optionally, you could open the document automatically (commented out to avoid interaction).
        // System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outputPath) { UseShellExecute = true });
    }
}
