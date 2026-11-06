using System;
using System.Runtime.InteropServices;

public class Program
{
    [STAThread]
    public static void Main()
    {
        // Attempt to create a COM (OLE Automation) object
        Type comType = Type.GetTypeFromProgID("WScript.Shell");
        object comObject = null;

        try
        {
            comObject = Activator.CreateInstance(comType);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to create COM object: {ex.Message}");
        }

        if (comObject != null)
        {
            Console.WriteLine("OLE object insertion successful: reference is not null.");
            Marshal.ReleaseComObject(comObject);
        }
        else
        {
            Console.WriteLine("OLE object insertion failed: reference is null.");
        }
    }
}
