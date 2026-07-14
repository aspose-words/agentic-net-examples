using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Math;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with some text.
        builder.Writeln("Sample document with equations:");

        // Insert a few equations using the deterministic EQ‑field bootstrap workflow.
        InsertEquation(builder, @"\f(1,2)");   // 1/2
        InsertEquation(builder, @"\r(3,x)");  // cube root of x
        InsertEquation(builder, @"\i \su(n=1,5,n)"); // integral with summation

        // Save the original document (optional, for reference).
        string originalPath = Path.Combine(Directory.GetCurrentDirectory(), "Original.docx");
        doc.Save(originalPath);

        // -------- Macro logic: toggle display mode of selected equations --------
        // For this example we treat all top‑level OfficeMath nodes (MathObjectType.OMathPara) as selected.
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        foreach (OfficeMath om in officeMathNodes)
        {
            if (om.MathObjectType != MathObjectType.OMathPara)
                continue; // Skip nested math objects.

            // Toggle between Inline and Display.
            if (om.DisplayType == OfficeMathDisplayType.Inline)
            {
                om.DisplayType = OfficeMathDisplayType.Display;
                // When in Display mode a justification can be set.
                om.Justification = OfficeMathJustification.Left;
            }
            else
            {
                om.DisplayType = OfficeMathDisplayType.Inline;
                // Justification cannot be set for Inline mode; leave it unchanged.
            }
        }

        // Save the modified document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Toggled.docx");
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }

    // Helper that inserts an EQ field, converts it to a real OfficeMath node, and removes the field.
    private static OfficeMath InsertEquation(DocumentBuilder builder, string eqArgument)
    {
        // Insert an EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
        // Write the equation argument into the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(eqArgument);
        // Return to the field start's parent (the paragraph).
        builder.MoveTo(field.Start.ParentNode);

        // Convert the field to an OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();
        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start.
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);
            // Remove the original field.
            field.Remove();
            // Move the builder after the inserted OfficeMath to continue building.
            builder.MoveTo(officeMath.NextSibling ?? officeMath);
            // Add a new paragraph after the equation for readability.
            builder.Writeln();
        }

        return officeMath;
    }
}
