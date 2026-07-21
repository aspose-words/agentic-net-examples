using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    public static void Main()
    {
        // Create a new document and builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert three simple equations using the deterministic EQ-field bootstrap workflow.
        for (int i = 0; i < 3; i++)
        {
            // Add a paragraph label.
            builder.Writeln($"Equation {i + 1}:");

            // Insert an EQ field.
            Field field = builder.InsertField(FieldType.FieldEquation, true);
            FieldEQ fieldEq = (FieldEQ)field;

            // Write a simple EQ argument.
            builder.MoveTo(fieldEq.Separator);
            builder.Write(@"\f(1,2)"); // simple fraction 1/2

            // Convert the field to a real OfficeMath node.
            OfficeMath officeMath = fieldEq.AsOfficeMath();
            if (officeMath != null)
            {
                // Insert the OfficeMath node before the field start.
                fieldEq.Start.ParentNode.InsertBefore(officeMath, fieldEq.Start);
                // Remove the original field.
                fieldEq.Remove();
                // Move the builder after the inserted OfficeMath node.
                builder.MoveTo(officeMath);
            }

            // Ensure the next equation starts on a new line.
            builder.Writeln();
        }

        // Save the initial document.
        string initialPath = "BulkUpdate.docx";
        doc.Save(initialPath);
        if (!File.Exists(initialPath))
            throw new InvalidOperationException($"Failed to create the initial document at '{initialPath}'.");

        // Perform bulk updates: set DisplayType to Display for all top‑level OfficeMath nodes.
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        foreach (OfficeMath om in officeMathNodes)
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
            {
                om.DisplayType = OfficeMathDisplayType.Display;
            }
        }

        // Validate that each OfficeMath node still has the expected MathObjectType.
        foreach (OfficeMath om in officeMathNodes)
        {
            if (om.MathObjectType != MathObjectType.OMathPara)
                throw new InvalidOperationException("An OfficeMath node does not have the expected MathObjectType OMathPara.");
        }

        // Save the updated document.
        string updatedPath = "BulkUpdateUpdated.docx";
        doc.Save(updatedPath);
        if (!File.Exists(updatedPath))
            throw new InvalidOperationException($"Failed to create the updated document at '{updatedPath}'.");

        // Indicate successful completion.
        Console.WriteLine("Bulk update completed successfully. Documents saved:");
        Console.WriteLine($"- {Path.GetFullPath(initialPath)}");
        Console.WriteLine($"- {Path.GetFullPath(updatedPath)}");
    }
}
