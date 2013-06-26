using System;
using System.Collections.Generic;
using System.Configuration;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SharpShell;
using SharpShell.Attributes;
using SharpShell.SharpContextMenu;
using TrelloNet;

namespace TrelloContextMenu
{
    [ComVisible(true)]
    [COMServerAssociation(AssociationType.ClassOfExtension, ".txt")]
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

    public class ContextMenuProvider
    {
        private readonly IEnumerable<ITestableContextMenu> contextMenus;

        public ContextMenuProvider(IEnumerable<ITestableContextMenu> contextMenus)
        {
            this.contextMenus = contextMenus;
        }

        public bool CanShowMenu(IEnumerable<string> selectedItemPaths, string folderPath)
        {
            return contextMenus.Any(x => x.CanShowMenu(selectedItemPaths, folderPath));
        }

        public ContextMenuStrip CreateMenu(Func<IEnumerable<string>> selectedItemPaths, string folderPath)
        {
            var ms = new ContextMenuStrip();

            var trelloMenu = new ToolStripMenuItem("Trello");
            contextMenus.ForEach(x =>
                {
                    if (x.CanShowMenu(selectedItemPaths(), folderPath))
                        trelloMenu.DropDownItems.Add(x.CreateMenuItem(selectedItemPaths, folderPath));
                });

            ms.Items.Add(trelloMenu);
            return ms;
        }
    }

    public class AddAsCardContextMenu : ITestableContextMenu
    {
        private TrelloItemProvider trello = TrelloItemProvider.Instance;

        public bool CanShowMenu(IEnumerable<string> selectedItemPaths, string folderPath)
        {
            return selectedItemPaths.Count(item => item.Contains(".txt")) == 1;
        }

        public ToolStripMenuItem CreateMenuItem(Func<IEnumerable<string>> selectedItemPaths, string folderPath)
        {
            var trelloItem = new ToolStripMenuItem("Add as card");

            trello.GetBoardNames().ForEach(
                boardName =>
                {
                    var board = new ToolStripMenuItem(boardName);
                    trello.GetListsForBoard(boardName).ForEach(
                        card => board.DropDownItems.Add(card));

                    board.DropDownItemClicked +=
                        (sender, args) => TrelloItemProvider.Instance.AddCard(boardName, args.ClickedItem.Text,
                                                            Path.GetFileNameWithoutExtension(selectedItemPaths().FirstOrDefault()),
                                                            File.ReadAllText(selectedItemPaths().FirstOrDefault()));

                    trelloItem.DropDownItems.Add(board);
                });

            return trelloItem;
        }
    }

    public class AddAsAttachmentContextMenu : ITestableContextMenu
    {
        private TrelloItemProvider trello = TrelloItemProvider.Instance;

        public bool CanShowMenu(IEnumerable<string> selectedItemPaths, string folderPath)
        {
            return true;
        }

        public ToolStripMenuItem CreateMenuItem(Func<IEnumerable<string>> selectedItemPaths, string folderPath)
        {
            var trelloItem = new ToolStripMenuItem("Add as attachment");

            trello.GetBoardNames().ForEach(
                boardName =>
                {
                    var board = new ToolStripMenuItem(boardName);
                    trello.GetListsForBoard(boardName).ForEach(
                        card => board.DropDownItems.Add(card));

                    board.DropDownItemClicked +=
                        (sender, args) => TrelloItemProvider.Instance.AddCard(boardName, args.ClickedItem.Text,
                                                            Path.GetFileNameWithoutExtension(selectedItemPaths().FirstOrDefault()),
                                                            File.ReadAllText(selectedItemPaths().FirstOrDefault()));

                    trelloItem.DropDownItems.Add(board);
                });

            return trelloItem;
        }
    }

    public class TrelloItemProvider
    {
        public static TrelloItemProvider Instance = new TrelloItemProvider();

        private TrelloItemProvider()
        {
            trello = new Trello("key");
            trello.Authorize("token");
        }

        private readonly Trello trello;

        public IEnumerable<string> GetBoardNames()
        {
            return trello.Boards.ForMe().Select(b => b.Name);
        }

        public IEnumerable<string> GetCardsForBoard(string boardName)
        {
            return trello.Cards.ForBoard(trello.Boards.ForMe().FirstOrDefault(b => b.Name == boardName)).Select(c => c.Name);
        }

        public IEnumerable<string> GetListsForBoard(string boardName)
        {
            return trello.Lists.ForBoard(trello.Boards.ForMe().FirstOrDefault(b => b.Name == boardName))
                      .Select(l => l.Name);
        }

        public void AddCard(string boardName, string listName, string cardName, string commentText = "")
        {
            var listid = new ListId(trello.Lists.ForBoard(trello.Boards.ForMe().First(b => b.Name == boardName))
                .First(l => l.Name == listName).GetListId());
            var card = trello.Cards.Add(cardName, listid);
            var cardid = new CardId(card.GetCardId());
            
            if(!string.IsNullOrWhiteSpace(commentText))
                trello.Cards.AddComment(cardid, commentText);
        }
    }

    public static class IKnowItsNotFunctionalButIWantItAnyways
    {
        public static void ForEach<T>(this IEnumerable<T> e, Action<T> action)
        {
            foreach (T item in e)
                action(item);
        }
    }
}
