using System;
using System.Collections.Generic;
using System.Configuration;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SharpShell;
using TrelloNet;

namespace TrelloContextMenu
{
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
        private readonly TrelloItemProvider trello = TrelloItemProvider.Instance;

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
                        list => board.DropDownItems.Add(list));

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
        private readonly TrelloItemProvider trello = TrelloItemProvider.Instance;

        public bool CanShowMenu(IEnumerable<string> selectedItemPaths, string folderPath)
        {
            return selectedItemPaths.All(f =>
                {
                    var attribs = File.GetAttributes(f);

                    // Verify that no selected items are directories (although attaching entire directories could be implemented)
                    if ((attribs & FileAttributes.Directory) != 0)
                        return false;

                    // All selected files must be less than Trello's attachment size limit (10mb)
                    return (new FileInfo(f)).Length < 10485760;
                });
        }

        public ToolStripMenuItem CreateMenuItem(Func<IEnumerable<string>> selectedItemPaths, string folderPath)
        {
            var trelloItem = new ToolStripMenuItem("Add as attachment");

            trello.GetBoardNames().ForEach(
                boardName =>
                {
                    var board = new ToolStripMenuItem(boardName);
                    trello.GetListsForBoard(boardName).ForEach(
                        list =>
                            {
                                var listItem = (ToolStripMenuItem) board.DropDownItems.Add(list);
                                trello.GetCardsForList(boardName, list).ForEach(
                                    card => listItem.DropDownItems.Add(card));

                                // Remove empty lists (todo: refactor this along with TrelloItemProvider)
                                if (listItem.DropDownItems.Count == 0)
                                    board.DropDownItems.Remove(listItem);
                                else
                                    listItem.DropDownItemClicked += (sender, args) =>
                                            selectedItemPaths().ForEach(file => 
                                                TrelloItemProvider.Instance.AttachToCard(boardName, list, args.ClickedItem.Text, file));
                            });

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
            trello = new Trello("[key]");
            trello.Authorize("[token]");
        }

        private readonly Trello trello;

        public IEnumerable<string> GetBoardNames()
        {
            return 
                trello.Boards.ForMe()
                .Select(b => b.Name);
        }

        public IEnumerable<string> GetCardsForBoard(string boardName)
        {
            return 
                trello.Cards.ForBoard(
                    trello.Boards.ForMe()
                    .FirstOrDefault(b => b.Name == boardName))
                .Select(c => c.Name);
        }

        public IEnumerable<string> GetListsForBoard(string boardName)
        {
            return 
                trello.Lists.ForBoard(
                    trello.Boards.ForMe()
                    .FirstOrDefault(b => b.Name == boardName))
                .Select(l => l.Name);
        }

        public IEnumerable<string> GetCardsForList(string boardName, string listName)
        {
            return
                trello.Cards.ForList(
                    trello.Lists.ForBoard(trello.Boards.ForMe()
                        .FirstOrDefault(b => b.Name == boardName))
                    .FirstOrDefault(l => l.Name == listName))
                .Select(c => c.Name);
        }

        public void AddCard(string boardName, string listName, string cardName, string commentText = "")
        {
            if(string.IsNullOrWhiteSpace(commentText))
                throw new Exception("Cannot add empty comment");

            var listid = new ListId(trello.Lists.ForBoard(trello.Boards.ForMe().First(b => b.Name == boardName))
                .First(l => l.Name == listName).GetListId());
            var card = trello.Cards.Add(cardName, listid);
            var cardid = new CardId(card.GetCardId());
            
            trello.Cards.AddComment(cardid, commentText);
        }

        public void AttachToCard(string boardName, string listName, string cardName, string fileToAttach)
        {
            if (!File.Exists(fileToAttach))
                throw new FileNotFoundException("File not found", fileToAttach);

            var listid = new ListId(trello.Lists.ForBoard(trello.Boards.ForMe().First(b => b.Name == boardName))
                .First(l => l.Name == listName).GetListId());
            var card = trello.Cards.ForList(listid)
                .First(c => c.Name == cardName).GetCardId();
            var cardid = new CardId(card);

            trello.Cards.AddAttachment(cardid, new FileAttachment(fileToAttach));
        }
    }
}
