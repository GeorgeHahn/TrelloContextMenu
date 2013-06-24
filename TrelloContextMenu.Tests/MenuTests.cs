using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FluentAssertions;
using SharpShell;
using SharpShell.SharpContextMenu;
using Xunit;

namespace TrelloContextMenu.Tests
{
    public class MenuTests
    {
        [Fact]
        public void WhenASingleFileIsSelected_ContextMenuShouldShow()
        {
            TestableContextMenu menu = new TestableContextMenu();
            var items = new[] { "blah.txt" };

            menu.CanShowMenu(items, "", "")
                .Should().BeTrue();
        }

        [Fact]
        public void WhenMultipleFilesAreSelected_ContextMenuShouldNotShow()
        {
            TestableContextMenu menu = new TestableContextMenu();
            var items = new[] { "blah.txt", "eh.txt" };

            menu.CanShowMenu(items, "", "")
                .Should().BeFalse();
        }

        [Fact]
        public void CreatedMenuHasOnlyOneItem()
        {
            TestableContextMenu menu = new TestableContextMenu();
            var items = new[] { "blah.txt" };

            menu.CreateMenu(() => items, "", "")
                .Items.Count
                .Should().Be(1);
        }
    }

    public class TrelloItemProviderTests
    {
        [Fact]
        public void BoardNames_AreReturned()
        {
            TrelloItemProvider.Instance.GetBoardNames()
                              .Count()
                              .Should().BeGreaterOrEqualTo(1);
        }

        [Fact]
        public void Cards_AreReturned()
        {
            var firstboard = TrelloItemProvider.Instance.GetBoardNames().First();
            TrelloItemProvider.Instance.GetCardsForBoard(firstboard)
                              .Count()
                              .Should().BeGreaterOrEqualTo(1);
        }

        [Fact]
        public void Lists_AreReturned()
        {
            var firstboard = TrelloItemProvider.Instance.GetBoardNames().First();
            TrelloItemProvider.Instance.GetListsForBoard(firstboard)
                              .Count()
                              .Should().BeGreaterOrEqualTo(1);
        }
    }
}
