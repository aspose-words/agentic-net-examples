using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class OfficeMathReplaceExample
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create a document with an initial equation (fraction 1/2).
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Paragraph before the equation:");

        // Insert EQ field and write the equation source.
        Field field = builder.InsertField(FieldType.FieldEquation, true);
        builder.MoveTo(field.Separator);
        builder.Write(@"\f(1,2)"); // Equation source string.
        field.Update(); // Convert the field to a real OfficeMath node.

        // Convert the field to OfficeMath.
        FieldEQ fieldEq = field as FieldEQ;
        OfficeMath originalMath = fieldEq?.AsOfficeMath();
        if (originalMath == null)
            throw new InvalidOperationException("Failed to create the initial OfficeMath node.");

        // Insert the OfficeMath node before the field start and remove the field.
        field.Start.ParentNode.InsertBefore(originalMath, field.Start);
        field.Remove();

        builder.Writeln(); // End of paragraph.

        // Save the original document.
        string originalPath = Path.Combine(outputDir, "Original.docx");
        doc.Save(originalPath);

        // 2. Locate the existing OfficeMath node to be replaced.
        OfficeMath existingMath = doc.GetChildNodes(NodeType.OfficeMath, true)
                                    .Cast<OfficeMath>()
                                    .FirstOrDefault(m => m.MathObjectType == MathObjectType.OMathPara);
        if (existingMath == null)
            throw new InvalidOperationException("No top‑level OfficeMath node found to replace.");

        // 3. Create a new OfficeMath node with the replacement equation (fraction 3/4).
        Document tempDoc = new Document();
        DocumentBuilder tempBuilder = new DocumentBuilder(tempDoc);

        Field tempField = tempBuilder.InsertField(FieldType.FieldEquation, true);
        tempBuilder.MoveTo(tempField.Separator);
        tempBuilder.Write(@"\f(3,4)"); // New equation source string.
        tempField.Update(); // Convert to OfficeMath.

        FieldEQ tempFieldEq = tempField as FieldEQ;
        OfficeMath newMath = tempFieldEq?.AsOfficeMath();
        if (newMath == null)
            throw new InvalidOperationException("Failed to create the replacement OfficeMath node.");

        // Insert the new OfficeMath before the temporary field start and remove the field.
        tempField.Start.ParentNode.InsertBefore(newMath, tempField.Start);
        tempField.Remove();

        // 4. Import the new OfficeMath node into the original document.
        NodeImporter importer = new NodeImporter(tempDoc, doc, ImportFormatMode.KeepSourceFormatting);
        Node importedNewMath = importer.ImportNode(newMath, true);
        if (importedNewMath == null)
            throw new InvalidOperationException("Failed to import the new OfficeMath node.");

        // Replace the existing OfficeMath with the new one.
        existingMath.ParentNode.InsertBefore(importedNewMath, existingMath);
        existingMath.Remove();

        // 5. Save the modified document.
        string replacedPath = Path.Combine(outputDir, "Replaced.docx");
        doc.Save(replacedPath);

        // 6. Validate that the document now contains the new equation.
        var mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true)
                           .Cast<OfficeMath>()
                           .Where(m => m.MathObjectType == MathObjectType.OMathPara)
                           .ToList();

        if (mathNodes.Count != 1)
            throw new InvalidOperationException("The document does not contain the expected number of top‑level OfficeMath nodes after replacement.");

        string mathText = mathNodes[0].GetText();
        if (!mathText.Contains("3") || !mathText.Contains("4"))
            throw new InvalidOperationException("The replacement equation was not found in the document.");

        Console.WriteLine("OfficeMath replacement completed. Files saved to: " + outputDir);
    }
}
