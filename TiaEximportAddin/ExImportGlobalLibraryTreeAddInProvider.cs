using Siemens.Engineering;
using Siemens.Engineering.AddIn;
using Siemens.Engineering.AddIn.Menu;

namespace TiaEximportAddin
{
    public class ExImportGlobalLibraryTreeAddInProvider : GlobalLibraryTreeAddInProvider
    {
        private readonly TiaPortal _tiaPortal;

        public ExImportGlobalLibraryTreeAddInProvider(TiaPortal tiaPortal)
        {
            _tiaPortal = tiaPortal;
        }

        protected override IEnumerable<ContextMenuAddIn> GetContextMenuAddIns()
        {
            yield return new AddIn(_tiaPortal);
        }
    }
}
