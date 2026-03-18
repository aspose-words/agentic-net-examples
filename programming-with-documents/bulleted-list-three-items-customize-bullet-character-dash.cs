using System;
using Aspose.Words;
using Aspose.Words.Lists;

class Program
{
    static void Main()
    {
        // Folder where the output document will be saved.
        string artifactsDir = @"C:\Artifacts\"; // Adjust path as needed.

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a list based on the default bullet template.
        List dashList = doc.Lists.Add(ListTemplate.BulletDefault);

        // Customize the first (and only) level of the list:
        // - Use a dash ("-") as the bullet character.
        // - Set the number style to Bullet so the NumberFormat is respected.
        // - Add a trailing space for readability (optional).
        ListLevel level = dashList.ListLevels[0];
        level.NumberStyle = NumberStyle.Bullet;
        level.NumberFormat = "-";
        level.TrailingCharacter = ListTrailingCharacter.Space;

        // Apply the custom list to the builder.
        builder.ListFormat.List = dashList;

        // Add three list items.
        builder.Writeln("First item");
        builder.Writeln("Second item");
        builder.Writeln("Third item");

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Save the document.
        doc.Save(artifactsDir + "CustomDashBulletList.docx");
    }
}
