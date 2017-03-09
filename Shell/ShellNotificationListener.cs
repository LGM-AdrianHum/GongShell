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
using System.Runtime.InteropServices;
using System.Windows.Forms;
using GongSolutions.Shell.Interop;

namespace GongSolutions.Shell
{
    /// <summary>
    ///     Listens for notifications of changes in the Windows Shell Namespace.
    /// </summary>
    public class ShellNotificationListener : Component
    {
        private readonly NotificationWindow _mWindow;

        /// <summary>
        ///     Initializes a new instance of the
        ///     <see cref="ShellNotificationListener" /> class.
        /// </summary>
        public ShellNotificationListener()
        {
            _mWindow = new NotificationWindow(this);
        }

        /// <summary>
        ///     Initializes a new instance of the
        ///     <see cref="ShellNotificationListener" /> class.
        /// </summary>
        public ShellNotificationListener(IContainer container)
        {
            container.Add(this);
            _mWindow = new NotificationWindow(this);
        }

        /// <summary>
        ///     Occurs when a drive is added.
        /// </summary>
        public event ShellItemEventHandler DriveAdded;

        /// <summary>
        ///     Occurs when a drive is removed.
        /// </summary>
        public event ShellItemEventHandler DriveRemoved;

        /// <summary>
        ///     Occurs when a folder is created.
        /// </summary>
        public event ShellItemEventHandler FolderCreated;

        /// <summary>
        ///     Occurs when a folder is deleted.
        /// </summary>
        public event ShellItemEventHandler FolderDeleted;

        /// <summary>
        ///     Occurs when a folder is renamed.
        /// </summary>
        public event ShellItemChangeEventHandler FolderRenamed;

        /// <summary>
        ///     Occurs when a folder's contents are updated.
        /// </summary>
        public event ShellItemEventHandler FolderUpdated;

        /// <summary>
        ///     Occurs when a non-folder item is created.
        /// </summary>
        public event ShellItemEventHandler ItemCreated;

        /// <summary>
        ///     Occurs when a non-folder item is deleted.
        /// </summary>
        public event ShellItemEventHandler ItemDeleted;

        /// <summary>
        ///     Occurs when a non-folder item is renamed.
        /// </summary>
        public event ShellItemChangeEventHandler ItemRenamed;

        /// <summary>
        ///     Occurs when a non-folder item is updated.
        /// </summary>
        public event ShellItemEventHandler ItemUpdated;

        /// <summary>
        ///     Occurs when the shared state for a folder changes.
        /// </summary>
        public event ShellItemEventHandler SharingChanged;

        /// <summary>
        ///     Overrides the <see cref="Component.Dispose(bool)" /> method.
        /// </summary>
        /// <param name="disposing" />
        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);
            _mWindow.Dispose();
        }

        private class NotificationWindow : Control
        {
            private const int WmShnotify = 0x401;

            private readonly uint _mNotifyId;
            private readonly ShellNotificationListener _mParent;

            public NotificationWindow(ShellNotificationListener parent)
            {
                var notify = new SHChangeNotifyEntry
                {
                    pidl = ShellItem.Desktop.Pidl,
                    fRecursive = true
                };
                _mNotifyId = Shell32.SHChangeNotifyRegister(Handle,
                    SHCNRF.InterruptLevel | SHCNRF.ShellLevel,
                    SHCNE.ALLEVENTS, WmShnotify, 1, ref notify);
                _mParent = parent;
            }

            protected override void Dispose(bool disposing)
            {
                base.Dispose(disposing);
                Shell32.SHChangeNotifyUnregister(_mNotifyId);
            }

            protected override void WndProc(ref Message m)
            {
                if (m.Msg == WmShnotify)
                {
                    var notify = (SHNOTIFYSTRUCT)
                        Marshal.PtrToStructure(m.WParam,
                            typeof(SHNOTIFYSTRUCT));

                    switch ((SHCNE) m.LParam)
                    {
                        case SHCNE.CREATE:
                            if (_mParent.ItemCreated != null)
                            {
                                var item = new ShellItem(notify.dwItem1);
                                _mParent.ItemCreated(_mParent,
                                    new ShellItemEventArgs(item));
                            }
                            break;

                        case SHCNE.DELETE:
                            if (_mParent.ItemDeleted != null)
                            {
                                var item = new ShellItem(notify.dwItem1);
                                _mParent.ItemDeleted(_mParent,
                                    new ShellItemEventArgs(item));
                            }
                            break;

                        case SHCNE.DRIVEADD:
                            if (_mParent.DriveAdded != null)
                            {
                                var item = new ShellItem(notify.dwItem1);
                                _mParent.DriveAdded(_mParent,
                                    new ShellItemEventArgs(item));
                            }
                            break;

                        case SHCNE.DRIVEREMOVED:
                            if (_mParent.DriveRemoved != null)
                            {
                                var item = new ShellItem(notify.dwItem1);
                                _mParent.DriveRemoved(_mParent,
                                    new ShellItemEventArgs(item));
                            }
                            break;

                        case SHCNE.MKDIR:
                            if (_mParent.FolderCreated != null)
                            {
                                var item = new ShellItem(notify.dwItem1);
                                _mParent.FolderCreated(_mParent,
                                    new ShellItemEventArgs(item));
                            }
                            break;

                        case SHCNE.RMDIR:
                            if (_mParent.FolderDeleted != null)
                            {
                                var item = new ShellItem(notify.dwItem1);
                                _mParent.FolderDeleted(_mParent,
                                    new ShellItemEventArgs(item));
                            }
                            break;

                        case SHCNE.UPDATEDIR:
                            if (_mParent.FolderUpdated != null)
                            {
                                var item = new ShellItem(notify.dwItem1);
                                _mParent.FolderUpdated(_mParent,
                                    new ShellItemEventArgs(item));
                            }
                            break;

                        case SHCNE.UPDATEITEM:
                            if (_mParent.ItemUpdated != null)
                            {
                                var item = new ShellItem(notify.dwItem1);
                                _mParent.ItemUpdated(_mParent,
                                    new ShellItemEventArgs(item));
                            }
                            break;

                        case SHCNE.RENAMEFOLDER:
                            if (_mParent.FolderRenamed != null)
                            {
                                var item1 = new ShellItem(notify.dwItem1);
                                var item2 = new ShellItem(notify.dwItem2);
                                _mParent.FolderRenamed(_mParent,
                                    new ShellItemChangeEventArgs(item1, item2));
                            }
                            break;

                        case SHCNE.RENAMEITEM:
                            if (_mParent.ItemRenamed != null)
                            {
                                var item1 = new ShellItem(notify.dwItem1);
                                var item2 = new ShellItem(notify.dwItem2);
                                _mParent.ItemRenamed(_mParent,
                                    new ShellItemChangeEventArgs(item1, item2));
                            }
                            break;

                        case SHCNE.NETSHARE:
                        case SHCNE.NETUNSHARE:
                            if (_mParent.SharingChanged != null)
                            {
                                var item = new ShellItem(notify.dwItem1);
                                _mParent.SharingChanged(_mParent,
                                    new ShellItemEventArgs(item));
                            }
                            break;
                    }
                }
                else
                {
                    base.WndProc(ref m);
                }
            }
        }
    }

    /// <summary>
    ///     Provides information of changes in the Windows Shell Namespace.
    /// </summary>
    public class ShellItemEventArgs : EventArgs
    {
        /// <summary>
        ///     Initializes a new instance of the
        ///     <see cref="ShellItemEventArgs" /> class.
        /// </summary>
        /// <param name="item">
        ///     The ShellItem that has changed.
        /// </param>
        public ShellItemEventArgs(ShellItem item)
        {
            Item = item;
        }

        /// <summary>
        ///     The ShellItem that has changed.
        /// </summary>
        public ShellItem Item { get; }
    }

    /// <summary>
    ///     Provides information of changes in the Windows Shell Namespace.
    /// </summary>
    public class ShellItemChangeEventArgs : EventArgs
    {
        /// <summary>
        ///     Initializes a new instance of the
        ///     <see cref="ShellItemChangeEventArgs" /> class.
        /// </summary>
        /// <param name="oldItem">
        ///     The ShellItem before the change
        /// </param>
        /// <param name="newItem">
        ///     The ShellItem after the change
        /// </param>
        public ShellItemChangeEventArgs(ShellItem oldItem,
            ShellItem newItem)
        {
            OldItem = oldItem;
            NewItem = newItem;
        }

        /// <summary>
        ///     The ShellItem before the change.
        /// </summary>
        public ShellItem OldItem { get; }

        /// <summary>
        ///     The ShellItem after the change.
        /// </summary>
        public ShellItem NewItem { get; }
    }

    /// <summary>
    ///     Represents the method that handles change notifications from
    ///     <see cref="ShellNotificationListener" />
    /// </summary>
    /// <param name="sender">
    ///     The source of the event.
    /// </param>
    /// <param name="e">
    ///     A <see cref="ShellItemEventArgs" /> that contains the data
    ///     for the event.
    /// </param>
    public delegate void ShellItemEventHandler(object sender,
        ShellItemEventArgs e);

    /// <summary>
    ///     Represents the method that handles change notifications from
    ///     <see cref="ShellNotificationListener" />
    /// </summary>
    /// <param name="sender">
    ///     The source of the event.
    /// </param>
    /// <param name="e">
    ///     A <see cref="ShellItemChangeEventArgs" /> that contains the data
    ///     for the event.
    /// </param>
    public delegate void ShellItemChangeEventHandler(object sender,
        ShellItemChangeEventArgs e);
}