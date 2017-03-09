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
using System.Windows.Forms;
using GongSolutions.Shell.Properties;

namespace GongSolutions.Shell
{
    internal partial class ShellItemBrowseForm : Form
    {
        public ShellItemBrowseForm()
        {
            InitializeComponent();

            var manager = new KnownFolderManager();

            SystemImageList.UseSystemImageList(knownFolderList);
            foreach (var knownFolder in manager)
            {
                try
                {
                    var shellItem = knownFolder.CreateShellItem();
                    var item = knownFolderList.Items.Add(knownFolder.Name,
                        shellItem.GetSystemImageListIndex(ShellIconType.LargeIcon, 0));

                    item.Tag = knownFolder;

                    if (item.Text == Resources.ShellItemBrowseForm_ShellItemBrowseForm_Personal)
                    {
                        item.Text = Resources.ShellItemBrowseForm_ShellItemBrowseForm_Personal__My_Documents_;
                        item.Group = knownFolderList.Groups["common"];
                    }
                    else if ((item.Text == Resources.ShellItemBrowseForm_ShellItemBrowseForm_Desktop) ||
                             (item.Text == Resources.ShellItemBrowseForm_ShellItemBrowseForm_Downloads) ||
                             (item.Text == Resources.ShellItemBrowseForm_ShellItemBrowseForm_MyComputerFolder))
                    {
                        item.Group = knownFolderList.Groups["common"];
                    }
                    else
                    {
                        item.Group = knownFolderList.Groups["all"];
                    }
                }
                catch (Exception)
                {
                    // ignored
                }
            }
        }

        public ShellItem SelectedItem { get; private set; }

        private void SaveSelection()
        {
            if (tabControl.SelectedTab == knownFoldersPage)
            {
                SelectedItem = ((KnownFolder) knownFolderList.SelectedItems[0].Tag)
                    .CreateShellItem();
            }
            else
            {
                SelectedItem = allFilesView.SelectedItems[0];
            }
        }

        private void knownFolderList_DoubleClick(object sender, EventArgs e)
        {
            if (knownFolderList.SelectedItems.Count > 0)
            {
                SaveSelection();
                DialogResult = DialogResult.OK;
            }
        }

        private void knownFolderList_SelectedIndexChanged(object sender, EventArgs e)
        {
            okButton.Enabled = knownFolderList.SelectedItems.Count > 0;
        }

        private void fileBrowseView_SelectionChanged(object sender, EventArgs e)
        {
            okButton.Enabled = allFilesView.SelectedItems.Length > 0;
        }

        private void tabControl_Selected(object sender, TabControlEventArgs e)
        {
            if (e.TabPage == knownFoldersPage)
            {
                okButton.Enabled = knownFolderList.SelectedItems.Count > 0;
            }
            else if (e.TabPage == allFilesPage)
            {
                okButton.Enabled = allFilesView.SelectedItems.Length > 0;
            }
        }

        private void okButton_Click(object sender, EventArgs e)
        {
            SaveSelection();
            DialogResult = DialogResult.OK;
        }
    }
}