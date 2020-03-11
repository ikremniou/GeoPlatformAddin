using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;

namespace ExcelAddin.Ribbon
{
    [ComVisible(true)]
    public class ExampleClass : ExcelRibbon
    {
        public void OnExample1ButtonPressed(IRibbonControl control)
        {
            // ...
        }
    }
}