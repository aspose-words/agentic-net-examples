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

        // Write some introductory text.
        builder.Write("Inline equation: ");

        // Insert an EQ field with a simple fraction argument.
        FieldEQ eqField = InsertFieldEQ(builder, @"\f(1,2)");

        // Ensure the field is up‑to‑date before conversion.
        eqField.Update();

        // Convert the EQ field to a real OfficeMath object.
        OfficeMath officeMath = eqField.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Replace the field with the OfficeMath node.
        eqField.Start.ParentNode.InsertBefore(officeMath, eqField.Start);
        eqField.Remove();

        // Set the equation to be displayed inline.
        officeMath.DisplayType = OfficeMathDisplayType.Inline;

        // Save the document.
        string outputPath = "OfficeMathInline.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }

    // Helper that follows the deterministic EQ‑field bootstrap workflow.
    private static FieldEQ InsertFieldEQ(DocumentBuilder builder, string argument)
    {
        // Insert an empty EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Move to the field separator and write the EQ argument (e.g., "\f(1,2)").
        builder.MoveTo(field.Separator);
        builder.Write(argument);

        // Return the builder to the paragraph that contains the field.
        builder.MoveTo(field.Start.ParentNode);
        return field;
    }
}
