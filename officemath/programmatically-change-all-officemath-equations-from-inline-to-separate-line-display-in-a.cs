using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Helper to insert an inline equation using the deterministic EQ-field bootstrap workflow.
        void InsertInlineEquation(string eqSwitch)
        {
            // Insert a field of type Equation.
            builder.InsertField(FieldType.FieldEquation, true);
            // Retrieve the inserted field.
            FieldStart fieldStart = (FieldStart)builder.CurrentParagraph.GetChildNodes(NodeType.FieldStart, true).Cast<FieldStart>().Last();
            Field field = fieldStart.GetField();
            FieldEQ fieldEq = field as FieldEQ;
            if (fieldEq == null)
                throw new InvalidOperationException("Failed to create FieldEQ.");

            // Write the EQ argument into the field separator.
            builder.MoveTo(fieldEq.Separator);
            builder.Write(eqSwitch);

            // Convert the field to a real OfficeMath node.
            OfficeMath officeMath = fieldEq.AsOfficeMath();
            if (officeMath == null)
                throw new InvalidOperationException("EQ field could not be converted to OfficeMath.");

            // Insert the OfficeMath node before the field start node.
            Node fieldStartNode = field.Start;
            fieldStartNode.ParentNode.InsertBefore(officeMath, fieldStartNode);

            // Remove the original field (including start, separator, and end nodes).
            field.Remove();
        }

        // Insert several paragraphs with inline equations.
        builder.Writeln("Report Title");
        builder.Writeln();

        builder.Writeln("Paragraph 1 with an inline equation:");
        InsertInlineEquation(@"\f(1,2)");
        builder.Writeln();

        builder.Writeln("Paragraph 2 with another inline equation:");
        InsertInlineEquation(@"\r(3,x)");
        builder.Writeln();

        builder.Writeln("Paragraph 3 with a third inline equation:");
        InsertInlineEquation(@"\f(5,6)");
        builder.Writeln();

        // Change all top‑level inline equations to display (separate line) mode.
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        int changedCount = 0;
        foreach (OfficeMath om in mathNodes)
        {
            if (om.MathObjectType == MathObjectType.OMathPara && om.DisplayType == OfficeMathDisplayType.Inline)
            {
                om.DisplayType = OfficeMathDisplayType.Display;
                changedCount++;
            }
        }

        // Save the modified document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ReportWithDisplayMath.docx");
        doc.Save(outputPath, SaveFormat.Docx);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);

        // Reload the document and verify that all equations are now in display mode.
        Document loadedDoc = new Document(outputPath);
        int displayCount = loadedDoc.GetChildNodes(NodeType.OfficeMath, true)
            .Cast<OfficeMath>()
            .Count(om => om.MathObjectType == MathObjectType.OMathPara && om.DisplayType == OfficeMathDisplayType.Display);

        int expectedCount = 3; // Number of equations we inserted.
        if (displayCount != expectedCount)
            throw new InvalidOperationException($"Expected {expectedCount} display equations, but found {displayCount}.");

        // Program completed successfully.
        Console.WriteLine($"Successfully changed {changedCount} equations to display mode. Output saved to: {outputPath}");
    }
}
