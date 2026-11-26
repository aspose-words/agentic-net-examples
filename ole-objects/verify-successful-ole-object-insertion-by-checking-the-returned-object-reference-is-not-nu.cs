using System;

public class Program
{
    public static void Main()
    {
        // Obtain the COM type using its ProgID
        Type comType = Type.GetTypeFromProgID("WScript.Shell");
        object oleObject = null;

        if (comType != null)
        {
            // Create an instance of the COM object (OLE object)
            oleObject = Activator.CreateInstance(comType);
        }

        // Verify that the object reference is not null
        bool insertionSuccessful = oleObject != null;

        Console.WriteLine(insertionSuccessful
            ? "OLE object insertion succeeded."
            : "OLE object insertion failed.");
    }
}
