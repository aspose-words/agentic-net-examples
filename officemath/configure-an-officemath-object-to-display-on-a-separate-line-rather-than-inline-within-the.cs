using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that introduces the equation.
        builder.Writeln("Below is an equation displayed on its own line:");

        // Insert an EQ field with a simple fraction.
        FieldEQ eqField = InsertFieldEQ(builder, @"\f(1,2)");

        // Ensure fields are up‑to‑date (not strictly required but safe).
        doc.UpdateFields();

        // Convert the EQ field to a real OfficeMath node.
        OfficeMath officeMath = eqField.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Replace the field with the OfficeMath node.
        eqField.Start.ParentNode.InsertBefore(officeMath, eqField.Start);
        eqField.Remove();

        // Apply display formatting only to top‑level equations.
        if (officeMath.MathObjectType == MathObjectType.OMathPara)
        {
            officeMath.DisplayType = OfficeMathDisplayType.Display;   // Show on its own line.
            officeMath.Justification = OfficeMathJustification.Left; // Left‑justify.
        }

        // Save the document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "OfficeMathDisplay.docx");
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);
    }

    // Helper that inserts an EQ field, writes the provided arguments, and starts a new paragraph.
    private static FieldEQ InsertFieldEQ(DocumentBuilder builder, string args)
    {
        // Insert the EQ field. The field code will be "EQ".
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, false);

        // Move to the field separator and write the EQ arguments (e.g., a fraction).
        builder.MoveTo(field.Separator);
        builder.Write(args);

        // Return the builder to the paragraph that contains the field and start a new line.
        builder.MoveTo(field.Start.ParentNode);
        builder.InsertParagraph();

        return field;
    }
}
