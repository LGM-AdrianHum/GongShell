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
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Design;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Forms;
using GongSolutions.Shell.Interop;
using Microsoft.Win32;
using IDataObject = System.Runtime.InteropServices.ComTypes.IDataObject;
using IDropTarget = GongSolutions.Shell.Interop.IDropTarget;

namespace GongSolutions.Shell
{
    /// <summary>
    ///     Provides a tree view of a computer's folders.
    /// </summary>
    /// <remarks>
    ///     <para>
    ///         The <see cref="ShellTreeView" /> control allows you to embed Windows
    ///         Explorer functionality in your Windows Forms applications. The
    ///         control provides a tree view of the computer's folders, as it would
    ///         appear in the left-hand pane in Explorer.
    ///     </para>
    /// </remarks>
    public class ShellTreeView : Control, IDropSource, IDropTarget
    {
        private readonly Timer _mScrollTimer = new Timer();
        private readonly ShellNotificationListener _mShellListener = new ShellNotificationListener();

        private readonly TreeView _mTreeView;
        private bool _mAllowDrop;
        private DragTarget _mDragTarget;
        private bool _mNavigating;
        private TreeNode _mRightClickNode;
        private ShellItem _mRootFolder = ShellItem.Desktop;
        private ScrollDirection _mScrollDirection = ScrollDirection.None;
        private ShellView _mShellView;
        private ShowHidden _mShowHidden = ShowHidden.System;

        /// <summary>
        ///     Initializes a new instance of the <see cref="ShellTreeView" /> class.
        /// </summary>
        public ShellTreeView()
        {
            _mTreeView = new TreeView
            {
                Dock = DockStyle.Fill,
                HideSelection = false,
                HotTracking = true,
                Parent = this,
                ShowRootLines = false
            };
            _mTreeView.AfterSelect += m_TreeView_AfterSelect;
            _mTreeView.BeforeExpand += m_TreeView_BeforeExpand;
            _mTreeView.ItemDrag += m_TreeView_ItemDrag;
            _mTreeView.MouseDown += m_TreeView_MouseDown;
            _mTreeView.MouseUp += m_TreeView_MouseUp;
            _mScrollTimer.Interval = 250;
            _mScrollTimer.Tick += m_ScrollTimer_Tick;
            Size = new Size(120, 100);
            SystemImageList.UseSystemImageList(_mTreeView);

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

            // Setting AllowDrop to true then false makes sure OleInitialize()
            // is called for the thread: it must be called before we can use
            // RegisterDragDrop. There is probably a neater way of doing this.
            _mTreeView.AllowDrop = true;
            _mTreeView.AllowDrop = false;

            CreateItems();
        }

        /// <summary>
        ///     Gets/sets a value indicating whether drag/drop operations are
        ///     allowed on the control.
        /// </summary>
        [DefaultValue(false)]
        public override bool AllowDrop
        {
            get { return _mAllowDrop; }
            set
            {
                if (value != _mAllowDrop)
                {
                    _mAllowDrop = value;

                    Marshal.ThrowExceptionForHR(
                        _mAllowDrop
                            ? Ole32.RegisterDragDrop(_mTreeView.Handle, this)
                            : Ole32.RevokeDragDrop(_mTreeView.Handle));
                }
            }
        }

        /// <summary>
        ///     Gets or sets a value indicating whether a tree node label takes on
        ///     the appearance of a hyperlink as the mouse pointer passes over it.
        /// </summary>
        [DefaultValue(true)]
        [Category("Appearance")]
        public bool HotTracking
        {
            get { return _mTreeView.HotTracking; }
            set { _mTreeView.HotTracking = value; }
        }

        /// <summary>
        ///     Gets or sets the root folder that is displayed in the
        ///     <see cref="ShellTreeView" />.
        /// </summary>
        [Category("Appearance")]
        public ShellItem RootFolder
        {
            get { return _mRootFolder; }
            set
            {
                _mRootFolder = value;
                CreateItems();
            }
        }

        /// <summary>
        ///     Gets/sets a <see cref="ShellView" /> whose navigation should be
        ///     controlled by the treeview.
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
        ///     Gets or sets the selected folder in the
        ///     <see cref="ShellTreeView" />.
        /// </summary>
        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        [Editor(typeof(ShellItemEditor), typeof(UITypeEditor))]
        public ShellItem SelectedFolder
        {
            get { return (ShellItem) _mTreeView.SelectedNode.Tag; }
            set { SelectItem(value); }
        }

        /// <summary>
        ///     Gets or sets a value indicating whether hidden folders should
        ///     be displayed in the tree.
        /// </summary>
        [DefaultValue(ShowHidden.System), Category("Appearance")]
        public ShowHidden ShowHidden
        {
            get { return _mShowHidden; }
            set
            {
                _mShowHidden = value;
                RefreshContents();
            }
        }

        #region Hidden Properties

        /// <summary>
        ///     This property does not apply to the <see cref="ShellTreeView" />
        ///     class.
        /// </summary>
        [Browsable(false)]
        [EditorBrowsable(EditorBrowsableState.Never)]
        public override string Text
        {
            get { return base.Text; }
            set { base.Text = value; }
        }

        #endregion

        /// <summary>
        ///     Refreses the contents of the <see cref="ShellTreeView" />.
        /// </summary>
        public void RefreshContents()
        {
            RefreshItem(_mTreeView.Nodes[0]);
        }

        /// <summary>
        ///     Occurs when the <see cref="SelectedFolder" /> property changes.
        /// </summary>
        public event EventHandler SelectionChanged;

        private void CreateItems()
        {
            _mTreeView.BeginUpdate();

            try
            {
                _mTreeView.Nodes.Clear();
                CreateItem(null, _mRootFolder);
                _mTreeView.Nodes[0].Expand();
                _mTreeView.SelectedNode = _mTreeView.Nodes[0];
            }
            finally
            {
                _mTreeView.EndUpdate();
            }
        }

        private void CreateItem(TreeNode parent, ShellItem folder)
        {
            var displayName = folder.DisplayName;

            var node = parent != null ? InsertNode(parent, folder, displayName) : _mTreeView.Nodes.Add(displayName);

            if (folder.HasSubFolders)
            {
                node.Nodes.Add("");
            }

            node.Tag = folder;
            SetNodeImage(node);
        }

        private void CreateChildren(TreeNode node)
        {
            if ((node.Nodes.Count == 1) && (node.Nodes[0].Tag == null))
            {
                var folder = (ShellItem) node.Tag;
                var e = GetFolderEnumerator(folder);

                node.Nodes.Clear();
                while (e.MoveNext())
                {
                    CreateItem(node, e.Current);
                }
            }
        }

        private void RefreshItem(TreeNode node)
        {
            var folder = (ShellItem) node.Tag;
            node.Text = folder.DisplayName;
            SetNodeImage(node);

            if (NodeHasChildren(node))
            {
                var e = GetFolderEnumerator(folder);
                var nodesToRemove = new ArrayList(node.Nodes);

                while (e.MoveNext())
                {
                    var childNode = FindItem(e.Current, node);

                    if (childNode != null)
                    {
                        RefreshItem(childNode);
                        nodesToRemove.Remove(childNode);
                    }
                    else
                    {
                        CreateItem(node, e.Current);
                    }
                }

                foreach (TreeNode n in nodesToRemove)
                {
                    n.Remove();
                }
            }
            else if (node.Nodes.Count == 0)
            {
                if (folder.HasSubFolders)
                {
                    node.Nodes.Add("");
                }
            }
        }

        private TreeNode InsertNode(TreeNode parent, ShellItem folder, string displayName)
        {
            var parentFolder = (ShellItem) parent.Tag;
            var folderRelPidl = Shell32.ILFindLastID(folder.Pidl);
            TreeNode result = null;

            foreach (TreeNode child in parent.Nodes)
            {
                var childFolder = (ShellItem) child.Tag;
                var childRelPidl = Shell32.ILFindLastID(childFolder.Pidl);
                var compare = parentFolder.GetIShellFolder().CompareIDs(0,
                    folderRelPidl, childRelPidl);

                if (compare < 0)
                {
                    result = parent.Nodes.Insert(child.Index, displayName);
                    break;
                }
            }

            return result ?? parent.Nodes.Add(displayName);
        }

        private bool ShouldShowHidden()
        {
            if (_mShowHidden != ShowHidden.System) return _mShowHidden == ShowHidden.True;
            var reg = Registry.CurrentUser.OpenSubKey(
                @"Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced");

            return (int?) reg?.GetValue("Hidden", 2) == 1;
        }

        private IEnumerator<ShellItem> GetFolderEnumerator(ShellItem folder)
        {
            var filter = SHCONTF.FOLDERS;
            if (ShouldShowHidden()) filter |= SHCONTF.INCLUDEHIDDEN;
            return folder.GetEnumerator(filter);
        }

        private void SetNodeImage(TreeNode node)
        {
            var itemInfo = new TVITEMW();
            var folder = (ShellItem) node.Tag;

            // We need to set the images for the item by sending a 
            // TVM_SETITEMW message, as we need to set the overlay images,
            // and the .Net TreeView API does not support overlays.
            itemInfo.mask = TVIF.TVIF_IMAGE | TVIF.TVIF_SELECTEDIMAGE |
                            TVIF.TVIF_STATE;
            itemInfo.hItem = node.Handle;
            itemInfo.iImage = folder.GetSystemImageListIndex(
                ShellIconType.SmallIcon, ShellIconFlags.OverlayIndex);
            itemInfo.iSelectedImage = folder.GetSystemImageListIndex(
                ShellIconType.SmallIcon, ShellIconFlags.OpenIcon);
            itemInfo.state = (TVIS) (itemInfo.iImage >> 16);
            itemInfo.stateMask = TVIS.TVIS_OVERLAYMASK;
            User32.SendMessage(_mTreeView.Handle, MSG.TVM_SETITEMW,
                0, ref itemInfo);
        }

        private void SelectItem(ShellItem value)
        {
            var node = _mTreeView.Nodes[0];
            var folder = (ShellItem) node.Tag;

            if (folder == value)
            {
                _mTreeView.SelectedNode = node;
            }
            else
            {
                SelectItem(node, value);
            }
        }

        private void SelectItem(TreeNode node, ShellItem value)
        {
            CreateChildren(node);

            foreach (TreeNode child in node.Nodes)
            {
                var folder = (ShellItem) child.Tag;

                if (folder == value)
                {
                    _mTreeView.SelectedNode = child;
                    child.EnsureVisible();
                    child.Expand();
                    return;
                }
                if (folder.IsParentOf(value))
                {
                    SelectItem(child, value);
                    return;
                }
            }
        }

        private TreeNode FindItem(ShellItem item, TreeNode parent)
        {
            if ((ShellItem) parent.Tag == item)
            {
                return parent;
            }

            foreach (TreeNode node in parent.Nodes)
            {
                if ((ShellItem) node.Tag == item)
                {
                    return node;
                }
                var found = FindItem(item, node);
                if (found != null) return found;
            }
            return null;
        }

        private bool NodeHasChildren(TreeNode node)
        {
            return (node.Nodes.Count > 0) && (node.Nodes[0].Tag != null);
        }

        private void ScrollTreeView(ScrollDirection direction)
        {
            User32.SendMessage(_mTreeView.Handle, MSG.WM_VSCROLL,
                (int) direction, 0);
        }

        private void CheckDragScroll(Point location)
        {
            var scrollArea = (int) (_mTreeView.Nodes[0].Bounds.Height*1.5);
            var scroll = ScrollDirection.None;

            if (location.Y < scrollArea)
            {
                scroll = ScrollDirection.Up;
            }
            else if (location.Y > _mTreeView.ClientRectangle.Height - scrollArea)
            {
                scroll = ScrollDirection.Down;
            }

            if (scroll != ScrollDirection.None)
            {
                if (_mScrollDirection == ScrollDirection.None)
                {
                    ScrollTreeView(scroll);
                    _mScrollTimer.Enabled = true;
                }
            }
            else
            {
                _mScrollTimer.Enabled = false;
            }

            _mScrollDirection = scroll;
        }

        // ReSharper disable once UnusedMember.Local
        private ShellItem[] ParseShellIdListArray(IDataObject pDataObj)
        {
            var result = new List<ShellItem>();
            var format = new FORMATETC();
            // ReSharper disable once RedundantAssignment
            var medium = new STGMEDIUM();

            format.cfFormat = (short) User32.RegisterClipboardFormat("Shell IDList Array");
            format.dwAspect = DVASPECT.DVASPECT_CONTENT;
            format.lindex = 0;
            format.ptd = IntPtr.Zero;
            format.tymed = TYMED.TYMED_HGLOBAL;

            pDataObj.GetData(ref format, out medium);
            Kernel32.GlobalLock(medium.unionmember);

            try
            {
                ShellItem parentFolder = null;
                var count = Marshal.ReadInt32(medium.unionmember);
                var offset = 4;

                for (var n = 0; n <= count; ++n)
                {
                    var pidlOffset = Marshal.ReadInt32(medium.unionmember, offset);
                    var pidlAddress = (int) medium.unionmember + pidlOffset;

                    if (n == 0)
                    {
                        parentFolder = new ShellItem(new IntPtr(pidlAddress));
                    }
                    else
                    {
                        result.Add(new ShellItem(parentFolder, new IntPtr(pidlAddress)));
                    }

                    offset += 4;
                }
            }
            finally
            {
                Marshal.FreeHGlobal(medium.unionmember);
            }

            return result.ToArray();
        }

        private bool ShouldSerializeRootFolder()
        {
            return _mRootFolder != ShellItem.Desktop;
        }

        private void m_TreeView_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if ((_mShellView != null) && !_mNavigating)
            {
                _mNavigating = true;
                try
                {
                    _mShellView.CurrentFolder = SelectedFolder;
                }
                catch (Exception)
                {
                    SelectedFolder = _mShellView.CurrentFolder;
                }
                finally
                {
                    _mNavigating = false;
                }
            }

            SelectionChanged?.Invoke(this, EventArgs.Empty);
        }

        private void m_TreeView_BeforeExpand(object sender, TreeViewCancelEventArgs e)
        {
            try
            {
                CreateChildren(e.Node);
            }
            catch (Exception)
            {
                e.Cancel = true;
            }
        }

        private void m_TreeView_ItemDrag(object sender, ItemDragEventArgs e)
        {
            var node = (TreeNode) e.Item;
            var folder = (ShellItem) node.Tag;
            DragDropEffects effect;

            Ole32.DoDragDrop(folder.GetIDataObject(), this,
                DragDropEffects.All, out effect);
        }

        private void m_TreeView_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                _mRightClickNode = _mTreeView.GetNodeAt(e.Location);
            }
        }

        private void m_TreeView_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                var node = _mTreeView.GetNodeAt(e.Location);

                if ((node != null) && (node == _mRightClickNode))
                {
                    var folder = (ShellItem) node.Tag;
                    new ShellContextMenu(folder).ShowContextMenu(_mTreeView, e.Location);
                }
            }
        }

        private void m_ScrollTimer_Tick(object sender, EventArgs e)
        {
            ScrollTreeView(_mScrollDirection);
        }

        private void m_ShellListener_ItemRenamed(object sender, ShellItemChangeEventArgs e)
        {
            var node = FindItem(e.OldItem, _mTreeView.Nodes[0]);
            if (node != null) RefreshItem(node);
        }

        private void m_ShellListener_ItemUpdated(object sender, ShellItemEventArgs e)
        {
            var parent = FindItem(e.Item.Parent, _mTreeView.Nodes[0]);
            if (parent != null) RefreshItem(parent);
        }

        private void m_ShellView_Navigated(object sender, EventArgs e)
        {
            if (!_mNavigating)
            {
                _mNavigating = true;
                SelectedFolder = _mShellView.CurrentFolder;
                _mNavigating = false;
            }
        }

        private enum ScrollDirection
        {
            None = -1,
            Up,
            Down
        }

        private class DragTarget : IDisposable
        {
            private readonly Timer _mDragExpandTimer;
            private readonly IDropTarget _mDropTarget;

            public DragTarget(TreeNode node,
                int keyState, Point pt,
                ref int effect)
            {
                Node = node;
                Node.BackColor = SystemColors.Highlight;
                Node.ForeColor = SystemColors.HighlightText;

                _mDragExpandTimer = new Timer {Interval = 1000};
                _mDragExpandTimer.Tick += m_DragExpandTimer_Tick;
                _mDragExpandTimer.Start();

                try
                {
                    _mDropTarget = Folder.GetIDropTarget(node.TreeView);
                    _mDropTarget.DragEnter(Data, keyState, pt, ref effect);
                }
                catch (Exception)
                {
                    // ignored
                }
            }

            // ReSharper disable once MemberCanBePrivate.Local
            public ShellItem Folder => (ShellItem) Node.Tag;

            public TreeNode Node { get; }

            public static IDataObject Data { private get; set; }

            public void Dispose()
            {
                Node.BackColor = Node.TreeView.BackColor;
                Node.ForeColor = Node.TreeView.ForeColor;
                _mDragExpandTimer.Dispose();

                _mDropTarget?.DragLeave();
            }

            public void DragOver(int keyState, Point pt, ref int effect)
            {
                if (_mDropTarget != null)
                {
                    _mDropTarget.DragOver(keyState, pt, ref effect);
                }
                else
                {
                    effect = 0;
                }
            }

            public void Drop(IDataObject data, int keyState,
                Point pt, ref int effect)
            {
                _mDropTarget.Drop(data, keyState, pt, ref effect);
            }

            private void m_DragExpandTimer_Tick(object sender, EventArgs e)
            {
                Node.Expand();
                _mDragExpandTimer.Stop();
            }
        }

        #region IDropSource Members

        HResult IDropSource.QueryContinueDrag(bool fEscapePressed, int grfKeyState)
        {
            if (fEscapePressed)
            {
                return HResult.DRAGDROP_S_CANCEL;
            }
            if ((grfKeyState & (int) (MK.MK_LBUTTON | MK.MK_RBUTTON)) == 0)
            {
                return HResult.DRAGDROP_S_DROP;
            }
            return HResult.S_OK;
        }

        HResult IDropSource.GiveFeedback(int dwEffect)
        {
            return HResult.DRAGDROP_S_USEDEFAULTCURSORS;
        }

        #endregion

        #region IDropTarget Members

        void IDropTarget.DragEnter(IDataObject pDataObj,
            int grfKeyState, Point pt,
            ref int pdwEffect)
        {
            var clientLocation = _mTreeView.PointToClient(pt);
            var node = _mTreeView.HitTest(clientLocation).Node;

            DragTarget.Data = pDataObj;
            _mTreeView.HideSelection = true;

            if (node != null)
            {
                _mDragTarget = new DragTarget(node, grfKeyState, pt,
                    ref pdwEffect);
            }
            else
            {
                pdwEffect = 0;
            }
        }

        void IDropTarget.DragOver(int grfKeyState, Point pt,
            ref int pdwEffect)
        {
            var clientLocation = _mTreeView.PointToClient(pt);
            var node = _mTreeView.HitTest(clientLocation).Node;

            CheckDragScroll(clientLocation);

            if (node != null)
            {
                if ((_mDragTarget == null) ||
                    (node != _mDragTarget.Node))
                {
                    _mDragTarget?.Dispose();

                    _mDragTarget = new DragTarget(node, grfKeyState,
                        pt, ref pdwEffect);
                }
                else
                {
                    _mDragTarget.DragOver(grfKeyState, pt, ref pdwEffect);
                }
            }
            else
            {
                pdwEffect = 0;
            }
        }

        void IDropTarget.DragLeave()
        {
            if (_mDragTarget != null)
            {
                _mDragTarget.Dispose();
                _mDragTarget = null;
            }
            _mTreeView.HideSelection = false;
        }

        void IDropTarget.Drop(IDataObject pDataObj,
            int grfKeyState, Point pt,
            ref int pdwEffect)
        {
            if (_mDragTarget != null)
            {
                _mDragTarget.Drop(pDataObj, grfKeyState, pt,
                    ref pdwEffect);
                _mDragTarget.Dispose();
                _mDragTarget = null;
            }
        }

        #endregion
    }

    /// Describes whether hidden files/folders should be displayed in a 
    /// control.
    public enum ShowHidden
    {
        /// <summary>
        ///     Hidden files/folders should not be displayed.
        /// </summary>
        False,

        /// <summary>
        ///     Hidden files/folders should be displayed.
        /// </summary>
        True,

        /// <summary>
        ///     The Windows Explorer "Show hidden files" setting should be used
        ///     to determine whether to show hidden files/folders.
        /// </summary>
        System
    }
}