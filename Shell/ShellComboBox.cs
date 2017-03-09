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
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Design;
using System.IO;
using System.Windows.Forms;
using GongSolutions.Shell.Interop;

namespace GongSolutions.Shell
{
    /// <summary>
    ///     Provides a drop-down list displaying the Windows Shell namespace.
    /// </summary>
    /// <remarks>
    ///     The <see cref="ShellComboBox" /> class displays a view of the Windows
    ///     Shell namespace in a drop-down list similar to that displayed in
    ///     a file open/save dialog.
    /// </remarks>
    public class ShellComboBox : Control
    {
        private static ShellItem _mComputer;

        private readonly ComboBox _mCombo = new ComboBox();
        private readonly TextBox _mEdit = new TextBox();
        private readonly ShellNotificationListener _mShellListener = new ShellNotificationListener();
        private bool _mChangingLocation;
        private bool _mCreatingItems;
        private bool _mEditable;
        private ShellItem _mRootFolder = ShellItem.Desktop;
        private bool _mSelectAll;
        private ShellItem _mSelectedFolder;
        private ShellView _mShellView;
        private bool _mShowFileSystemPath;

        /// <summary>
        ///     Initializes a new instance of the <see cref="ShellComboBox" /> class.
        /// </summary>
        public ShellComboBox()
        {
            _mCombo.Dock = DockStyle.Fill;
            _mCombo.DrawMode = DrawMode.OwnerDrawFixed;
            _mCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            _mCombo.DropDownHeight = 300;
            _mCombo.ItemHeight = SystemInformation.SmallIconSize.Height + 1;
            _mCombo.Parent = this;
            _mCombo.Click += m_Combo_Click;
            _mCombo.DrawItem += m_Combo_DrawItem;
            _mCombo.SelectedIndexChanged += m_Combo_SelectedIndexChanged;

            _mEdit.Anchor = AnchorStyles.Left | AnchorStyles.Top |
                            AnchorStyles.Right | AnchorStyles.Bottom;
            _mEdit.BorderStyle = BorderStyle.None;
            _mEdit.Left = 8 + SystemInformation.SmallIconSize.Width;
            _mEdit.Top = 4;
            _mEdit.Width = Width - _mEdit.Left - 3 - SystemInformation.VerticalScrollBarWidth;
            _mEdit.Parent = this;
            _mEdit.Visible = false;
            _mEdit.GotFocus += m_Edit_GotFocus;
            _mEdit.LostFocus += m_Edit_LostFocus;
            _mEdit.KeyDown += m_Edit_KeyDown;
            _mEdit.MouseDown += m_Edit_MouseDown;
            _mEdit.BringToFront();

            _mShellListener.DriveAdded += m_ShellListener_ItemUpdated;
            _mShellListener.DriveRemoved += m_ShellListener_ItemUpdated;
            _mShellListener.FolderCreated += m_ShellListener_ItemUpdated;
            _mShellListener.FolderDeleted += m_ShellListener_ItemUpdated;
            _mShellListener.FolderRenamed += m_ShellListener_ItemRenamed;
            _mShellListener.FolderUpdated += m_ShellListener_ItemUpdated;
            _mShellListener.ItemCreated += m_ShellListener_ItemUpdated;
            _mShellListener.ItemDeleted += m_ShellListener_ItemUpdated;
            _mShellListener.ItemRenamed += m_ShellListener_ItemRenamed;
            _mShellListener.ItemUpdated += m_ShellListener_ItemUpdated;
            _mShellListener.SharingChanged += m_ShellListener_ItemUpdated;

            _mSelectedFolder = ShellItem.Desktop;
            _mEdit.Text = GetEditString();

            if (_mComputer == null)
            {
                _mComputer = new ShellItem(Environment.SpecialFolder.MyComputer);
            }

            CreateItems();
        }

        /// <summary>
        ///     Gets/sets a value indicating whether the combo box is editable.
        /// </summary>
        [DefaultValue(false)]
        public bool Editable
        {
            get { return _mEditable; }
            set { _mEdit.Visible = _mEditable = value; }
        }

        /// <summary>
        ///     Gets/sets a value indicating whether the full file system path
        ///     should be displayed in the main portion of the control.
        /// </summary>
        [DefaultValue(false)]
        public bool ShowFileSystemPath
        {
            get { return _mShowFileSystemPath; }
            set
            {
                _mShowFileSystemPath = value;
                _mCombo.Invalidate();
            }
        }

        /// <summary>
        ///     Gets/sets the folder that the <see cref="ShellComboBox" /> should
        ///     display as the root folder.
        /// </summary>
        [Editor(typeof(ShellItemEditor), typeof(UITypeEditor))]
        public ShellItem RootFolder
        {
            get { return _mRootFolder; }
            set
            {
                _mRootFolder = value;
                if (!_mRootFolder.IsParentOf(_mSelectedFolder))
                {
                    _mSelectedFolder = _mRootFolder;
                }
                CreateItems();
            }
        }

        /// <summary>
        ///     Gets/sets the folder currently selected in the
        ///     <see cref="ShellComboBox" />.
        /// </summary>
        [Editor(typeof(ShellItemEditor), typeof(UITypeEditor))]
        public ShellItem SelectedFolder
        {
            get { return _mSelectedFolder; }
            set
            {
                if (_mSelectedFolder != value)
                {
                    _mSelectedFolder = value;
                    CreateItems();
                    _mEdit.Text = GetEditString();
                    NavigateShellView();
                    OnChanged();
                }
            }
        }

        /// <summary>
        ///     Gets/sets a <see cref="ShellView" /> whose navigation should be
        ///     controlled by the combo box.
        /// </summary>
        [DefaultValue(null), Category("Behaviour")]
        public ShellView ShellView
        {
            get { return _mShellView; }
            set
            {
                if (_mShellView != null)
                {
                    _mShellView.Navigated -= m_ShellView_Navigated;
                }

                _mShellView = value;

                if (_mShellView != null)
                {
                    _mShellView.Navigated += m_ShellView_Navigated;
                    m_ShellView_Navigated(_mShellView, EventArgs.Empty);
                }
            }
        }

        /// <summary>
        ///     Occurs when the <see cref="ShellComboBox" />'s
        ///     <see cref="SelectedFolder" /> property changes.
        /// </summary>
        public event EventHandler Changed;

        /// <summary>
        ///     Occurs when the <see cref="ShellComboBox" /> control wants to know
        ///     if it should include a folder in its view.
        /// </summary>
        /// <remarks>
        ///     This event allows the folders displayed in the
        ///     <see cref="ShellComboBox" /> control to be filtered.
        /// </remarks>
        public event FilterItemEventHandler FilterItem;

        internal bool ShouldSerializeRootFolder()
        {
            return _mRootFolder != ShellItem.Desktop;
        }

        internal bool ShouldSerializeSelectedFolder()
        {
            return _mSelectedFolder != ShellItem.Desktop;
        }

        private void CreateItems()
        {
            if (!_mCreatingItems)
            {
                try
                {
                    _mCreatingItems = true;
                    _mCombo.Items.Clear();
                    CreateItem(_mRootFolder, 0);
                }
                finally
                {
                    _mCreatingItems = false;
                }
            }
        }

        private void CreateItems(ShellItem folder, int indent)
        {
            var e = folder.GetEnumerator(
                SHCONTF.FOLDERS | SHCONTF.INCLUDEHIDDEN);

            while (e.MoveNext())
            {
                if (ShouldCreateItem(e.Current))
                {
                    CreateItem(e.Current, indent);
                }
            }
        }

        private void CreateItem(ShellItem folder, int indent)
        {
            var index = _mCombo.Items.Add(new ComboItem(folder, indent));

            if (folder == _mSelectedFolder)
            {
                _mCombo.SelectedIndex = index;
            }

            if (ShouldCreateChildren(folder))
            {
                CreateItems(folder, indent + 1);
            }
        }

        private bool ShouldCreateItem(ShellItem folder)
        {
            var e = new FilterItemEventArgs(folder);
            // ReSharper disable once UnusedVariable
            var myComputer = new ShellItem(Environment.SpecialFolder.MyComputer);

            e.Include = false;

            if (ShellItem.Desktop.IsImmediateParentOf(folder) ||
                _mComputer.IsImmediateParentOf(folder))
            {
                e.Include = folder.IsFileSystemAncestor;
            }
            else if ((folder == _mSelectedFolder) ||
                     folder.IsParentOf(_mSelectedFolder))
            {
                e.Include = true;
            }

            FilterItem?.Invoke(this, e);

            return e.Include;
        }

        private bool ShouldCreateChildren(ShellItem folder)
        {
            return (folder == _mComputer) ||
                   (folder == ShellItem.Desktop) ||
                   folder.IsParentOf(_mSelectedFolder);
        }

        private string GetEditString()
        {
            if (_mShowFileSystemPath && _mSelectedFolder.IsFileSystem)
            {
                return _mSelectedFolder.FileSystemPath;
            }
            return _mSelectedFolder.DisplayName;
        }

        private void NavigateShellView()
        {
            if ((_mShellView != null) && !_mChangingLocation)
            {
                try
                {
                    _mChangingLocation = true;
                    _mShellView.Navigate(_mSelectedFolder);
                }
                catch (Exception)
                {
                    SelectedFolder = _mShellView.CurrentFolder;
                }
                finally
                {
                    _mChangingLocation = false;
                }
            }
        }

        private void OnChanged()
        {
            Changed?.Invoke(this, EventArgs.Empty);
        }

        private void m_Combo_Click(object sender, EventArgs e)
        {
            OnClick(e);
        }

        private void m_Combo_DrawItem(object sender, DrawItemEventArgs e)
        {
            var iconWidth = SystemInformation.SmallIconSize.Width;
            var indent = (e.State & DrawItemState.ComboBoxEdit) == 0
                ? iconWidth/2
                : 0;

            if (e.Index != -1)
            {
                string display;
                var item = (ComboItem) _mCombo.Items[e.Index];
                var textColor = SystemColors.WindowText;

                if ((e.State & DrawItemState.ComboBoxEdit) != 0)
                {
                    // Don't draw the folder location in the edit box when
                    // the control is Editable as the edit control will
                    // take care of that.
                    display = _mEditable ? string.Empty : GetEditString();
                }
                else
                {
                    display = item.Folder.DisplayName;
                }

                SizeF size = TextRenderer.MeasureText(display, _mCombo.Font);

                var textRect = new Rectangle(
                    e.Bounds.Left + iconWidth + item.Indent*indent + 3,
                    e.Bounds.Y, (int) size.Width, e.Bounds.Height);
                var textOffset = (int) ((e.Bounds.Height - size.Height)/2);

                // If the text is being drawin in the main combo box edit area,
                // draw the text 1 pixel higher - this is how it looks in Windows.
                if ((e.State & DrawItemState.ComboBoxEdit) != 0)
                {
                    textOffset -= 1;
                }

                if ((e.State & DrawItemState.Selected) != 0)
                {
                    e.Graphics.FillRectangle(SystemBrushes.Highlight, textRect);
                    textColor = SystemColors.HighlightText;
                }
                else
                {
                    e.DrawBackground();
                }

                if ((e.State & DrawItemState.Focus) != 0)
                {
                    ControlPaint.DrawFocusRectangle(e.Graphics, textRect);
                }

                SystemImageList.DrawSmallImage(e.Graphics,
                    new Point(e.Bounds.Left + item.Indent*indent,
                        e.Bounds.Top),
                    item.Folder.GetSystemImageListIndex(ShellIconType.SmallIcon,
                        ShellIconFlags.OverlayIndex),
                    (e.State & DrawItemState.Selected) != 0);
                TextRenderer.DrawText(e.Graphics, display, _mCombo.Font,
                    new Point(textRect.Left, textRect.Top + textOffset),
                    textColor);
            }
        }

        private void m_Combo_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!_mCreatingItems)
            {
                SelectedFolder = ((ComboItem) _mCombo.SelectedItem).Folder;
            }
        }

        private void m_Edit_GotFocus(object sender, EventArgs e)
        {
            _mEdit.SelectAll();
            _mSelectAll = true;
        }

        private void m_Edit_LostFocus(object sender, EventArgs e)
        {
            _mSelectAll = false;
        }

        private void m_Edit_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                var path = _mEdit.Text;

                if ((path == string.Empty) ||
                    (string.Compare(path, "Desktop", StringComparison.OrdinalIgnoreCase) == 0))
                {
                    SelectedFolder = ShellItem.Desktop;
                    return;
                }

                if (Directory.Exists(path))
                {
                    SelectedFolder = new ShellItem(path);
                    return;
                }

                path = Path.Combine(_mSelectedFolder.FileSystemPath, path);

                if (Directory.Exists(path))
                {
                    SelectedFolder = new ShellItem(path);
                }
            }
        }

        private void m_Edit_MouseDown(object sender, MouseEventArgs e)
        {
            if (_mSelectAll)
            {
                _mEdit.SelectAll();
                _mSelectAll = false;
            }
            else
            {
                _mEdit.SelectionStart = _mEdit.Text.Length;
            }
        }

        private void m_ShellView_Navigated(object sender, EventArgs e)
        {
            if (!_mChangingLocation)
            {
                try
                {
                    _mChangingLocation = true;
                    SelectedFolder = _mShellView.CurrentFolder;
                    OnChanged();
                }
                finally
                {
                    _mChangingLocation = false;
                }
            }
        }

        private void m_ShellListener_ItemRenamed(object sender, ShellItemChangeEventArgs e)
        {
            CreateItems();
        }

        private void m_ShellListener_ItemUpdated(object sender, ShellItemEventArgs e)
        {
            CreateItems();
        }

        private class ComboItem
        {
            public readonly ShellItem Folder;
            public readonly int Indent;

            public ComboItem(ShellItem folder, int indent)
            {
                Folder = folder;
                Indent = indent;
            }
        }
    }
}