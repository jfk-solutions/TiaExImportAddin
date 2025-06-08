using Siemens.Engineering;
using Siemens.Engineering.AddIn.Menu;
using Siemens.Engineering.Hmi.RuntimeScripting;
using Siemens.Engineering.Hmi.Screen;
using Siemens.Engineering.Hmi.Tag;
using Siemens.Engineering.Hmi.TextGraphicList;
using Siemens.Engineering.SW.Blocks;
using Siemens.Engineering.SW.Tags;
using Siemens.Engineering.SW.Types;
using Siemens.Engineering.SW.WatchAndForceTables;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace TiaEximportAddin
{
    public class AddIn : ContextMenuAddIn
    {
        private readonly TiaPortal _tiaPortal;

        public AddIn(TiaPortal tiaPortal) : base("Clipboard")
        {
            _tiaPortal = tiaPortal;
        }

        protected override void BuildContextMenuItems(ContextMenuAddInRoot addInRootSubmenu)
        {
            addInRootSubmenu.Items.AddActionItem<IEngineeringObject>("Export", ExportToClipboard, MenuStatusForExport);
            addInRootSubmenu.Items.AddActionItem<IEngineeringObject>("Import", ImportFromClipboard, MenuStatusForImport);
        }

        public void ExportToClipboard(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider)
        {
            using var val = _tiaPortal.ExclusiveAccess("Exporting...");
            var list = menuSelectionProvider.GetSelection<IEngineeringObject>().ToList();

            var exportOptions = ExportOptions.None;

            try
            {
                foreach (var eo in list)
                {
                    var exportFile = Path.GetTempFileName().Replace(".tmp", ".xml");
                    var exportFileInfo = new FileInfo(exportFile);
                    switch (eo)
                    {
                        case PlcBlock plcBlock:
                            plcBlock.Export(exportFileInfo, exportOptions);
                            break;
                        case PlcType plcType:
                            plcType.Export(exportFileInfo, exportOptions);
                            break;
                        case PlcTagTable plcTagTable:
                            plcTagTable.Export(exportFileInfo, exportOptions);
                            break;
                        case PlcWatchTable plcWatchTable:
                            plcWatchTable.Export(exportFileInfo, exportOptions);
                            break;
                        case TagTable tagtable:
                            tagtable.Export(exportFileInfo, exportOptions);
                            break;
                        case TextList textList:
                            textList.Export(exportFileInfo, exportOptions);
                            break;
                    }
                    var text = File.ReadAllText(exportFile);
                    File.Delete(exportFile);
                    Clipboard.SetText(text);
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        public MenuStatus MenuStatusForExport(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider)
        {
            foreach (object item in menuSelectionProvider.GetSelection())
            {
                if (
                    item is PlcBlock ||
                    item is PlcTagTable ||
                    item is PlcForceTable ||
                    item is PlcWatchTable ||
                    item is PlcType ||
                    item is TagTable ||
                    item is TagFolder ||
                    item is VBScript ||
                    item is VBScriptFolder ||
                    item is GraphicList ||
                    item is GraphicListComposition ||
                    item is TextList ||
                    item is TextListComposition)
                {
                    return MenuStatus.Enabled;
                }
            }
            return MenuStatus.Disabled;
        }

        public void ImportFromClipboard(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider)
        {
            try
            {
                var xml = Clipboard.GetText();

                if (string.IsNullOrEmpty(xml))
                    return;

                var regEx = new Regex(@"<Engineering version=""(.*)"" ");
                xml = regEx.Replace(xml, @"<Engineering version=""V19"" ");

                using var val = _tiaPortal.ExclusiveAccess("Exporting...");
                var list = menuSelectionProvider.GetSelection<IEngineeringObject>().ToList();

                var importOptions = ImportOptions.None;

                foreach (var eo in list)
                {
                    var importFile = Path.GetTempFileName().Replace(".tmp", ".xml");
                    File.WriteAllText(importFile, xml);
                    var importFileInfo = new FileInfo(importFile);
                    switch (eo)
                    {
                        case PlcBlockSystemGroup plcBlockSystemGroup:
                            plcBlockSystemGroup.Blocks.Import(importFileInfo, importOptions);
                            break;
                        case PlcBlockGroup plcBlockGroup:
                            plcBlockGroup.Blocks.Import(importFileInfo, importOptions);
                            break;
                        case PlcTagTableGroup plcTagTableGroup:
                            plcTagTableGroup.TagTables.Import(importFileInfo, importOptions);
                            break;
                        case PlcWatchAndForceTableGroup plcWatchAndForceTableGroup:
                            try
                            {
                                plcWatchAndForceTableGroup.WatchTables.Import(importFileInfo, importOptions);
                            }
                            catch { }
                            try
                            {
                                plcWatchAndForceTableGroup.ForceTables.Import(importFileInfo, importOptions);
                            }
                            catch { }
                            break;
                        case PlcTypeGroup plcTypeGroup:
                            plcTypeGroup.Types.Import(importFileInfo, importOptions);
                            break;
                    }
                    File.Delete(importFile);
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        public MenuStatus MenuStatusForImport(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider)
        {
            foreach (object item in menuSelectionProvider.GetSelection())
            {
                if (item is PlcBlockSystemGroup ||
                    item is PlcBlockGroup ||
                    item is PlcTagTableGroup ||
                    item is PlcTypeGroup ||
                    item is PlcWatchAndForceTableGroup ||
                    item is ScreenFolder || 
                    item is ScreenTemplateFolder || 
                    item is ScreenPopupFolder || 
                    item is ScreenSlideinSystemFolder || 
                    item is ScreenGlobalElements || 
                    item is ScreenOverview || 
                    item is TagFolder || 
                    item is VBScriptFolder)
                {
                    return MenuStatus.Enabled;
                }
            }
            return MenuStatus.Disabled;
        }
    }
}
