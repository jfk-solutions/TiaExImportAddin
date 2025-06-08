using Siemens.Engineering;
using Siemens.Engineering.AddIn;
using Siemens.Engineering.AddIn.Menu;

namespace TiaEximportAddin
{
    public class ExImportProjectTreeAddInProvider : ProjectTreeAddInProvider
    {
        private readonly TiaPortal _tiaPortal;

        public ExImportProjectTreeAddInProvider(TiaPortal tiaPortal)
        {
            _tiaPortal = tiaPortal;
        }

        protected override IEnumerable<ContextMenuAddIn> GetContextMenuAddIns()
        {
            yield return new AddIn(_tiaPortal);
        }
    }
}
