using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace OfficeCliAgent.Word
{
    /// <summary>Ribbon (XML) with one toggle button on the Home tab. Ribbon.xml must be an Embedded Resource.</summary>
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;

        public string GetCustomUI(string ribbonID)
        {
            var xml = GetResourceText("OfficeCliAgent.Word.Ribbon.xml");
            if (xml == null)
                OfficeCliAgent.Core.Log.Write("ribbon: embedded resource 'OfficeCliAgent.Word.Ribbon.xml' not found. Set Ribbon.xml Build Action = Embedded Resource. Available: " +
                    string.Join(", ", Assembly.GetExecutingAssembly().GetManifestResourceNames()));
            return xml;
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI) => _ribbon = ribbonUI;

        public void OnTogglePane(Office.IRibbonControl control, bool pressed) => Globals.ThisAddIn.TogglePane();

        public bool GetPanePressed(Office.IRibbonControl control) => Globals.ThisAddIn.IsPaneVisible;

        public string GetLabel(Office.IRibbonControl control) => OfficeCliAgent.Core.UiLanguage.T("エージェント", "Agent");

        public string GetSupertip(Office.IRibbonControl control) =>
            OfficeCliAgent.Core.UiLanguage.T("Copilot エージェントのチャットを開閉します。", "Show or hide the Copilot agent chat.");

        public void Invalidate() => _ribbon?.InvalidateControl("OfficeCliAgentToggle");

        private static string GetResourceText(string resourceName)
        {
            var asm = Assembly.GetExecutingAssembly();
            foreach (var name in asm.GetManifestResourceNames())
            {
                if (string.Compare(resourceName, name, StringComparison.OrdinalIgnoreCase) != 0) continue;
                using (var reader = new StreamReader(asm.GetManifestResourceStream(name)))
                    return reader.ReadToEnd();
            }
            return null;
        }
    }
}
