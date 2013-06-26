using System.Runtime.InteropServices;
using System.Windows.Forms;
using SharpShell.Attributes;
using SharpShell.SharpContextMenu;

namespace TrelloContextMenu
{
    [ComVisible(true)]
    [COMServerAssociation(AssociationType.AllFiles)]
    public class UntestableContextMenu : SharpContextMenu
    {
        private readonly ContextMenuProvider provider;

        public UntestableContextMenu()
            : base()
        {
            provider = new ContextMenuProvider(
                new ITestableContextMenu[]
                    {
                        new AddAsAttachmentContextMenu(), 
                        new AddAsCardContextMenu()
                    });
        }

        protected override bool CanShowMenu()
        {
            return provider.CanShowMenu(SelectedItemPaths, FolderPath);
        }

        protected override ContextMenuStrip CreateMenu()
        {
            return provider.CreateMenu(() => SelectedItemPaths, FolderPath);
        }
    }
}