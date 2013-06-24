﻿using System;
using System.Collections.Generic;
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
        private readonly TestableContextMenu menu = new TestableContextMenu();

        protected override bool CanShowMenu()
        {
            return menu.CanShowMenu(SelectedItemPaths, FolderPath, DisplayName);
        }

        protected override ContextMenuStrip CreateMenu()
        {
            return menu.CreateMenu(() => SelectedItemPaths, FolderPath, DisplayName);
        }
    }

    public class TestableContextMenu
    {
        private TrelloItemProvider trello = TrelloItemProvider.Instance;
        
        public bool CanShowMenu(IEnumerable<string> selectedItemPaths, string folderPath, string displayName)
        {
            return selectedItemPaths.Count(item => item.Contains(".txt")) == 1;
        }

        public ContextMenuStrip CreateMenu(Func<IEnumerable<string>> selectedItemPaths, string folderPath, string displayName)
        {
            var ms = new ContextMenuStrip();
            var trelloItem = new ToolStripMenuItem("Trello");

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

            ms.Items.Add(trelloItem);
            return ms;
        }
    }

    public class TrelloItemProvider
    {
        public static TrelloItemProvider Instance = new TrelloItemProvider();

        private TrelloItemProvider()
        {
            trello = new Trello("[key]");
            var s = trello.GetAuthorizationUrl("Test", Scope.ReadWrite, Expiration.Never);

            trello.Authorize("[token]");
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
