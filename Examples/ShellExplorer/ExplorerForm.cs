// GongSolutions.Shell - A Windows Shell library for .Net.
// Copyright (C) 2007-2009 Steven J. Kirk
//
// This program is free software; you can redistribute it and/or
// modify it under the terms of the GNU Lesser General Public
// License as published by the Free Software Foundation; either 
// version 2 of the License, or (at your option) any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU Lesser General Public License for more details.
//
// You should have received a copy of the GNU Lesser General Public 
// License along with this program; if not, write to the Free 
// Software Foundation, Inc., 51 Franklin Street, Fifth Floor,  
// Boston, MA 2110-1301, USA.
//

using System;
using System.Collections.Generic;
using System.Windows.Forms;
using GongSolutions.Shell;

namespace ShellExplorer
{
    public partial class ShellExplorer : Form
    {
        private ShellContextMenu m_ContextMenu;

        public ShellExplorer()
        {
            InitializeComponent();
        }

        protected override void WndProc(ref Message m)
        {
            if ((m_ContextMenu == null) || !m_ContextMenu.HandleMenuMessage(ref m))
            {
                base.WndProc(ref m);
            }
        }

        private void shellView_Navigated(object sender, EventArgs e)
        {
            backButton.Enabled = shellView.CanNavigateBack;
            forwardButton.Enabled = shellView.CanNavigateForward;
            upButton.Enabled = shellView.CanNavigateParent;
        }

        private void fileMenu_Popup(object sender, EventArgs e)
        {
            var selectedItems = shellView.SelectedItems;

            if (selectedItems.Length > 0)
            {
                m_ContextMenu = new ShellContextMenu(selectedItems);
            }
            else
            {
                m_ContextMenu = new ShellContextMenu(treeView.SelectedFolder);
            }

            m_ContextMenu.Populate(fileMenu);
        }

        private void refreshMenu_Click(object sender, EventArgs e)
        {
            shellView.RefreshContents();
            treeView.RefreshContents();
        }

        private void toolBar_ButtonClick(object sender, ToolBarButtonClickEventArgs e)
        {
            if (e.Button == backButton)
            {
                shellView.NavigateBack();
            }
            else if (e.Button == forwardButton)
            {
                shellView.NavigateForward();
            }
            else if (e.Button == upButton)
            {
                shellView.NavigateParent();
            }
        }

        private void backButton_Popup(object sender, EventArgs e)
        {
            var items = new List<MenuItem>();

            backButtonMenu.MenuItems.Clear();
            foreach (var f in shellView.History.HistoryBack)
            {
                var item = new MenuItem(f.DisplayName);
                item.Tag = f;
                item.Click += backButtonMenuItem_Click;
                items.Insert(0, item);
            }

            backButtonMenu.MenuItems.AddRange(items.ToArray());
        }

        private void forwardButton_Popup(object sender, EventArgs e)
        {
            forwardButtonMenu.MenuItems.Clear();
            foreach (var f in shellView.History.HistoryForward)
            {
                var item = new MenuItem(f.DisplayName);
                item.Tag = f;
                item.Click += forwardButtonMenuItem_Click;
                forwardButtonMenu.MenuItems.Add(item);
            }
        }

        private void backButtonMenuItem_Click(object sender, EventArgs e)
        {
            var item = (MenuItem) sender;
            var folder = (ShellItem) item.Tag;
            shellView.NavigateBack(folder);
        }

        private void forwardButtonMenuItem_Click(object sender, EventArgs e)
        {
            var item = (MenuItem) sender;
            var folder = (ShellItem) item.Tag;
            shellView.NavigateForward(folder);
        }

        private void ShellExplorer_ResizeEnd(object sender, EventArgs e)
        {
            var calculatedWidth = shellView.Width - shellView.GetColumnWidth(1)
                                  - shellView.GetColumnWidth(2) - shellView.GetColumnWidth(3) - 25;
            shellView.SetColumnWidth(0, calculatedWidth);
        }
    }
}