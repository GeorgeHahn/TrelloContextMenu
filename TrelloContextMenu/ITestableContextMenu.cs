using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace TrelloContextMenu
{
    public interface ITestableContextMenu
    {
        bool CanShowMenu(IEnumerable<string> selectedItemPaths, string folderPath);
        ToolStripMenuItem CreateMenuItem(Func<IEnumerable<string>> selectedItemPaths, string folderPath);
    }
}