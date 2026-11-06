using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;   // Required for OfficeMath

public class OfficeMathInlineExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an EQ field that will later be converted to a real OfficeMath object.
        Field field = builder.InsertField(FieldType.FieldEquation, true);
        FieldEQ eqField = field as FieldEQ;
        if (eqField == null)
            throw new InvalidOperationException("Inserted field is not a FieldEQ.");

        // Move to the field separator and write a simple EQ argument (fraction).
        builder.MoveTo(eqField.Separator);
        builder.Write(@"\f(1,2)"); // Simple fraction equation.

        // Update the field so that Aspose.Words parses the EQ argument.
        eqField.Update();

        // Convert the EQ field to an OfficeMath node.
        OfficeMath officeMath = eqField.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("EQ field could not be converted to OfficeMath.");

        // Insert the OfficeMath node before the field start node.
        Node fieldStart = eqField.Start;
        fieldStart.ParentNode.InsertBefore(officeMath, fieldStart);

        // Remove the original EQ field, leaving only the real OfficeMath node.
        eqField.Remove();

        // Save the document.
        string outputPath = "OfficeMathInline.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);

        // Validate that the document contains at least one OfficeMath node.
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        if (mathNodes.Count == 0)
            throw new InvalidOperationException("No OfficeMath nodes were found in the saved document.");

        // Example completed successfully.
        Console.WriteLine($"Document saved to '{outputPath}' with {mathNodes.Count} OfficeMath node(s).");
    }
}
