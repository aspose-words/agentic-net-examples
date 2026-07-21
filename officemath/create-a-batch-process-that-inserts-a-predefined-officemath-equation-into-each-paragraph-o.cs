using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class OfficeMathBatchInsert
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample paragraphs to the document.
        builder.Writeln("First paragraph.");
        builder.Writeln("Second paragraph.");
        builder.Writeln("Third paragraph.");

        // Prepare a list of all paragraphs to avoid collection modification issues.
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true)
                            .Cast<Paragraph>()
                            .ToList();

        // Predefined EQ field argument (simple fraction 1/2).
        const string equationArgs = @"\f(1,2)";

        foreach (var paragraph in paragraphs)
        {
            // Move the builder cursor to the start of the current paragraph.
            builder.MoveTo(paragraph);

            // Insert an EQ field (the field will be updated immediately).
            Field field = builder.InsertField(FieldType.FieldEquation, true);
            FieldEQ fieldEQ = field as FieldEQ;
            if (fieldEQ == null)
                continue; // Safety check.

            // Write the equation arguments into the field separator.
            builder.MoveTo(fieldEQ.Separator);
            builder.Write(equationArgs);

            // Convert the EQ field to a real OfficeMath object.
            OfficeMath officeMath = fieldEQ.AsOfficeMath();
            if (officeMath != null)
            {
                // Insert the OfficeMath node before the field start.
                fieldEQ.Start.ParentNode.InsertBefore(officeMath, fieldEQ.Start);
                // Remove the original EQ field from the document.
                fieldEQ.Remove();
            }
        }

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "BatchOfficeMath.docx");
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
