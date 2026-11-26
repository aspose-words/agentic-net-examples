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

        // Insert first equation using the deterministic EQ‑field bootstrap workflow.
        InsertEquation(builder, @"\f(1,2)");

        // Insert second equation.
        InsertEquation(builder, @"\r(3,x)");

        // Set all top‑level OfficeMath equations to inline display mode.
        foreach (OfficeMath om in doc.GetChildNodes(NodeType.OfficeMath, true))
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
                om.DisplayType = OfficeMathDisplayType.Inline;
        }

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath, SaveFormat.Docx);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        // Reload and verify that the display type is inline.
        Document loaded = new Document(outputPath);
        foreach (OfficeMath om in loaded.GetChildNodes(NodeType.OfficeMath, true))
        {
            if (om.MathObjectType == MathObjectType.OMathPara &&
                om.DisplayType != OfficeMathDisplayType.Inline)
                throw new Exception("An equation was not set to inline display mode.");
        }
    }

    // Helper method to insert an equation via EQ field and convert it to OfficeMath.
    private static void InsertEquation(DocumentBuilder builder, string eqArgument)
    {
        // Insert an empty EQ field.
        Field field = builder.InsertField(FieldType.FieldEquation, true);
        FieldEQ fieldEq = field as FieldEQ;
        if (fieldEq == null)
            throw new InvalidOperationException("Failed to create FieldEQ.");

        // Move to the field separator and write the EQ argument.
        // A leading space is required after the field code.
        builder.MoveTo(fieldEq.Separator);
        builder.Write(" " + eqArgument);

        // Update the field so that Aspose.Words can convert it to OfficeMath.
        field.Update();

        // Convert the field to an OfficeMath object.
        OfficeMath officeMath = fieldEq.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("EQ field could not be converted to OfficeMath.");

        // Insert the OfficeMath node before the field start.
        Node fieldStart = fieldEq.Start;
        fieldStart.ParentNode.InsertBefore(officeMath, fieldStart);

        // Remove the original field.
        fieldEq.Remove();

        // Move the builder to the inserted equation for further content.
        builder.MoveTo(officeMath);
        builder.Writeln(); // Add a paragraph break after the equation.
    }
}
