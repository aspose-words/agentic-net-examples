using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class OfficeMathBatchInsert
{
    // Predefined EQ field argument that creates a simple fraction: \f(1,2)
    private const string EquationArgs = @"\f(1,2)";

    public static void Main()
    {
        // Create a sample document with several paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("First paragraph.");
        builder.Writeln("Second paragraph.");
        builder.Writeln("Third paragraph.");

        // Get the original set of paragraphs before any modifications.
        NodeCollection originalParagraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        int originalParagraphCount = originalParagraphs.Count;

        // Insert the predefined equation into every original paragraph.
        foreach (Paragraph para in originalParagraphs)
        {
            InsertEquationIntoParagraph(doc, para, EquationArgs);
        }

        // Save the result.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OutputWithEquations.docx");
        doc.Save(outputPath);

        // Validation: ensure the file exists and the expected number of equations were added.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not created.");

        // Count only top‑level OfficeMath nodes (MathObjectType == OMathPara).
        int equationCount = 0;
        foreach (OfficeMath om in doc.GetChildNodes(NodeType.OfficeMath, true))
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
                equationCount++;
        }

        if (equationCount != originalParagraphCount)
            throw new InvalidOperationException($"Expected {originalParagraphCount} equations, but found {equationCount}.");
    }

    // Inserts an OfficeMath equation into the specified paragraph using the EQ‑field bootstrap workflow.
    private static void InsertEquationIntoParagraph(Document doc, Paragraph paragraph, string eqArgs)
    {
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveTo(paragraph);

        // Insert an EQ field.
        Field eqField = builder.InsertField(FieldType.FieldEquation, true);
        FieldEQ fieldEQ = eqField as FieldEQ;
        if (fieldEQ == null)
            throw new InvalidOperationException("Failed to create an EQ field.");

        // Write the equation arguments into the field separator.
        if (fieldEQ.Separator == null)
            throw new InvalidOperationException("EQ field separator is missing.");

        builder.MoveTo(fieldEQ.Separator);
        builder.Write(eqArgs);

        // Ensure the field code is up‑to‑date before conversion.
        fieldEQ.Update();

        // Convert the field to a real OfficeMath object.
        OfficeMath officeMath = fieldEQ.AsOfficeMath();

        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Insert the OfficeMath node before the field start node.
        fieldEQ.Start.ParentNode.InsertBefore(officeMath, fieldEQ.Start);

        // Apply display formatting only to top‑level equations.
        if (officeMath.MathObjectType == MathObjectType.OMathPara)
        {
            officeMath.DisplayType = OfficeMathDisplayType.Display;
            officeMath.Justification = OfficeMathJustification.Left;
        }

        // Remove the original field, leaving only the OfficeMath node.
        fieldEQ.Remove();
    }
}
