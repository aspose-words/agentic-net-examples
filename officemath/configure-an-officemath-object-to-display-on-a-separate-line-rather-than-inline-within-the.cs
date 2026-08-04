using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add an introductory paragraph.
        builder.Writeln("Below is an equation displayed on its own line:");

        // Insert an EQ field that will be converted to a real OfficeMath object.
        FieldEQ eqField = InsertFieldEQ(builder, @"\f(1,2)");

        // Ensure the field is up‑to‑date so that AsOfficeMath can parse it.
        eqField.Update();

        // Convert the EQ field to an OfficeMath node.
        OfficeMath officeMath = eqField.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Replace the field with the OfficeMath node.
        eqField.Start.ParentNode.InsertBefore(officeMath, eqField.Start);
        eqField.Remove();

        // Verify that we have a top‑level equation (OMathPara).
        if (officeMath.MathObjectType != MathObjectType.OMathPara)
            throw new InvalidOperationException("The created OfficeMath is not a top‑level equation.");

        // Set the equation to display on its own line and left‑justify it.
        officeMath.DisplayType = OfficeMathDisplayType.Display;
        officeMath.Justification = OfficeMathJustification.Left;

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OfficeMathDisplay.docx");
        doc.Save(outputPath, SaveFormat.Docx);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);

        // Reload the document to ensure the saved settings are persisted.
        Document loadedDoc = new Document(outputPath);
        OfficeMath savedMath = (OfficeMath)loadedDoc.GetChild(NodeType.OfficeMath, 0, true);
        if (savedMath == null || savedMath.DisplayType != OfficeMathDisplayType.Display)
            throw new InvalidOperationException("The OfficeMath display type was not set to Display.");
    }

    // Helper that inserts an EQ field, writes the arguments, adds a following paragraph, and returns the field.
    private static FieldEQ InsertFieldEQ(DocumentBuilder builder, string args)
    {
        // Insert an empty EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Move to the field separator and write the EQ arguments.
        builder.MoveTo(field.Separator);
        builder.Write(args);

        // Return the builder to the paragraph that contains the field.
        builder.MoveTo(field.Start.ParentNode);
        // Insert a new paragraph after the field so the equation stands alone.
        builder.InsertParagraph();

        return field;
    }
}
