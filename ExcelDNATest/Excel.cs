using System;
using System.Windows.Forms;
using ExcelDna.Integration;

/// <summary>
/// The class here implements the ExcelDna.Integration.IExcelAddIn interface.
/// This allows the add-in to run code at start-up and shutdown.
/// </summary>
public class MainAddIn : IExcelAddIn
{
    public void AutoOpen()
    {
        MessageBox.Show("Hello, AutoOpen");
    }
    public void AutoClose()
    {
    }
}

/// <summary>
/// The class here is just for convenience; these methods wrap calls to VBA UDFs.
/// </summary>
static class Helper
{
    public static void SetScreen()
    {
        XlCall.Excel(XlCall.xlUDF, "main.xlsm!SetScreen");
    }
    public static void SetColor(ExcelReference er, int r, int g, int b)
    {
        XlCall.Excel(XlCall.xlUDF, "main.xlsm!SetColor", er, r, g, b);
    }
}

/// <summary>
/// The class here is also for convenience; The LoadCSharp macro calls Init().
/// </summary>
public class MainFunction
{
    static ExcelReference DrawArea = new ExcelReference(1, 72, 1, 128);
    public static void Init()
    {
        Helper.SetScreen();
        Helper.SetColor(DrawArea, 0, 0, 0);
    }
}