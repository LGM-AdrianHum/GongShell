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
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using GongSolutions.Shell.Interop;

namespace GongSolutions.Shell
{
    /// <summary>
    ///     Provides support for displaying the context menu of a shell item.
    /// </summary>
    /// <remarks>
    ///     <para>
    ///         Use this class to display a context menu for a shell item, either
    ///         as a popup menu, or as a main menu.
    ///     </para>
    ///     <para>
    ///         To display a popup menu, simply call <see cref="ShowContextMenu" />
    ///         with the parent control and the position at which the menu should
    ///         be shown.
    ///     </para>
    ///     <para>
    ///         To display a shell context menu in a Form's main menu, call the
    ///         <see cref="Populate" /> method to populate the menu. In addition,
    ///         you must intercept a number of special messages that will be sent
    ///         to the menu's parent form. To do this, you must override
    ///         <see cref="Form.WndProc" /> like so:
    ///     </para>
    ///     <code>
    ///     protected override void WndProc(ref Message m) {
    ///         if ((m_ContextMenu == null) || (!m_ContextMenu.HandleMenuMessage(ref m))) {
    ///             base.WndProc(ref m);
    ///         }
    ///     }
    /// </code>
    ///     <para>
    ///         Where m_ContextMenu is the <see cref="ShellContextMenu" /> being shown.
    ///     </para>
    ///     Standard menu commands can also be invoked from this class, for
    ///     example <see cref="InvokeDelete" /> and <see cref="InvokeRename" />.
    /// </remarks>
    public class ShellContextMenu
    {
        private const int MCmdFirst = 0x8000;
        private IContextMenu2 _mComInterface2;
        private IContextMenu3 _mComInterface3;

        private MessageWindow _mMessageWindow;

        /// <summary>
        ///     Initialises a new instance of the <see cref="ShellContextMenu" />
        ///     class.
        /// </summary>
        /// <param name="item">
        ///     The item to which the context menu should refer.
        /// </param>
        public ShellContextMenu(ShellItem item)
        {
            Initialize(new[] {item});
        }

        /// <summary>
        ///     Initialises a new instance of the <see cref="ShellContextMenu" />
        ///     class.
        /// </summary>
        /// <param name="items">
        ///     The items to which the context menu should refer.
        /// </param>
        public ShellContextMenu(ShellItem[] items)
        {
            Initialize(items);
        }

        /// <summary>
        ///     Gets the underlying COM <see cref="IContextMenu" /> interface.
        /// </summary>
        public IContextMenu ComInterface { get; set; }

        /// <summary>
        ///     Handles context menu messages when the <see cref="ShellContextMenu" />
        ///     is displayed on a Form's main menu bar.
        /// </summary>
        /// <remarks>
        ///     <para>
        ///         To display a shell context menu in a Form's main menu, call the
        ///         <see cref="Populate" /> method to populate the menu with the shell
        ///         item's menu items. In addition, you must intercept a number of
        ///         special messages that will be sent to the menu's parent form. To
        ///         do this, you must override <see cref="Form.WndProc" /> like so:
        ///     </para>
        ///     <code>
        ///     protected override void WndProc(ref Message m) {
        ///         if ((m_ContextMenu == null) || (!m_ContextMenu.HandleMenuMessage(ref m))) {
        ///             base.WndProc(ref m);
        ///         }
        ///     }
        /// </code>
        ///     <para>
        ///         Where m_ContextMenu is the <see cref="ShellContextMenu" /> being shown.
        ///     </para>
        /// </remarks>
        /// <param name="m">
        ///     The message to handle.
        /// </param>
        /// <returns>
        ///     <see langword="true" /> if the message was a Shell Context Menu
        ///     message, <see langword="false" /> if not. If the method returns false,
        ///     then the message should be passed down to the base class's
        ///     <see cref="Form.WndProc" /> method.
        /// </returns>
        public bool HandleMenuMessage(ref Message m)
        {
            if ((m.Msg == (int) MSG.WM_COMMAND) && ((int) m.WParam >= MCmdFirst))
            {
                InvokeCommand((int) m.WParam - MCmdFirst);
                return true;
            }
            if (_mComInterface3 != null)
            {
                IntPtr result;
                if (_mComInterface3.HandleMenuMsg2(m.Msg, m.WParam, m.LParam,
                    out result) == HResult.S_OK)
                {
                    m.Result = result;
                    return true;
                }
            }
            else
            {
                if (_mComInterface2?.HandleMenuMsg(m.Msg, m.WParam, m.LParam)
                    == HResult.S_OK)
                {
                    m.Result = IntPtr.Zero;
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        ///     Invokes the Copy command on the shell item(s).
        /// </summary>
        public void InvokeCopy()
        {
            InvokeVerb("copy");
        }

        /// <summary>
        ///     Invokes the Copy command on the shell item(s).
        /// </summary>
        public void InvokeCut()
        {
            InvokeVerb("cut");
        }

        /// <summary>
        ///     Invokes the Delete command on the shell item(s).
        /// </summary>
        public void InvokeDelete()
        {
            try
            {
                InvokeVerb("delete");
            }
            catch (COMException e)
            {
                // Ignore the exception raised when the user cancels
                // a delete operation.
                if (e.ErrorCode != unchecked((int) 0x800704C7) &&
                    e.ErrorCode != unchecked((int) 0x80270000))
                {
                    throw;
                }
            }
        }

        /// <summary>
        ///     Invokes the Paste command on the shell item(s).
        /// </summary>
        public void InvokePaste()
        {
            InvokeVerb("paste");
        }

        /// <summary>
        ///     Invokes the Rename command on the shell item.
        /// </summary>
        public void InvokeRename()
        {
            InvokeVerb("rename");
        }

        /// <summary>
        ///     Invokes the specified verb on the shell item(s).
        /// </summary>
        public void InvokeVerb(string verb)
        {
            var invoke = new CMINVOKECOMMANDINFO();
            invoke.cbSize = Marshal.SizeOf(invoke);
            invoke.lpVerb = verb;
            ComInterface.InvokeCommand(ref invoke);
        }

        /// <summary>
        ///     Populates a <see cref="Menu" /> with the context menu items for
        ///     a shell item.
        /// </summary>
        /// <remarks>
        ///     If this method is being used to populate a Form's main menu
        ///     then you need to call <see cref="HandleMenuMessage" /> in the
        ///     Form's message handler.
        /// </remarks>
        /// <param name="menu">
        ///     The menu to populate.
        /// </param>
        public void Populate(Menu menu)
        {
            RemoveShellMenuItems(menu);
            ComInterface.QueryContextMenu(menu.Handle, 0,
                MCmdFirst, int.MaxValue, CMF.EXPLORE);
        }

        /// <summary>
        ///     Shows a context menu for a shell item.
        /// </summary>
        /// <param name="control">
        ///     The parent control.
        /// </param>
        /// <param name="pos">
        ///     The position on <paramref name="control" /> that the menu
        ///     should be displayed at.
        /// </param>
        public void ShowContextMenu(Control control, Point pos)
        {
            using (var menu = new ContextMenu())
            {
                pos = control.PointToScreen(pos);
                Populate(menu);
                var command = User32.TrackPopupMenuEx(menu.Handle,
                    TPM.TPM_RETURNCMD, pos.X, pos.Y, _mMessageWindow.Handle,
                    IntPtr.Zero);
                if (command > 0)
                {
                    InvokeCommand(command - MCmdFirst);
                }
            }
        }

        private void Initialize(ShellItem[] items)
        {
            var pidls = new IntPtr[items.Length];
            ShellItem parent = null;

            for (var n = 0; n < items.Length; ++n)
            {
                pidls[n] = Shell32.ILFindLastID(items[n].Pidl);

                if (parent == null)
                {
                    parent = items[n] == ShellItem.Desktop ? ShellItem.Desktop : items[n].Parent;
                }
                else
                {
                    if (items[n].Parent != parent)
                    {
                        throw new Exception("All shell items must have the same parent");
                    }
                }
            }

            if (parent != null)
            {
                IntPtr result;
                parent.GetIShellFolder().GetUIObjectOf(IntPtr.Zero,
                    (uint) pidls.Length, pidls,
                    typeof(IContextMenu).GUID, 0, out result);

                ComInterface = (IContextMenu)
                    Marshal.GetTypedObjectForIUnknown(result,
                        typeof(IContextMenu));


                _mComInterface2 = ComInterface as IContextMenu2;
                _mComInterface3 = ComInterface as IContextMenu3;
                _mMessageWindow = new MessageWindow(this);
            }
        }

        private void InvokeCommand(int index)
        {
            const int swShownormal = 1;
            var invoke = new CMINVOKECOMMANDINFO_ByIndex();
            invoke.cbSize = Marshal.SizeOf(invoke);
            invoke.iVerb = index;
            invoke.nShow = swShownormal;
            _mComInterface2.InvokeCommand(ref invoke);
        }

        private void TagManagedMenuItems(Menu menu, int tag)
        {
            var info = new MENUINFO();

            info.cbSize = Marshal.SizeOf(info);
            info.fMask = MIM.MIM_MENUDATA;
            info.dwMenuData = tag;

            foreach (MenuItem item in menu.MenuItems)
            {
                User32.SetMenuInfo(item.Handle, ref info);
            }
        }

        private void RemoveShellMenuItems(Menu menu)
        {
            const int tag = 0xAB;
            var remove = new List<int>();
            var count = User32.GetMenuItemCount(menu.Handle);
            var menuInfo = new MENUINFO();
            var itemInfo = new MENUITEMINFO();

            menuInfo.cbSize = Marshal.SizeOf(menuInfo);
            menuInfo.fMask = MIM.MIM_MENUDATA;
            itemInfo.cbSize = Marshal.SizeOf(itemInfo);
            itemInfo.fMask = MIIM.MIIM_ID | MIIM.MIIM_SUBMENU;

            // First, tag the managed menu items with an arbitary 
            // value (0xAB).
            TagManagedMenuItems(menu, tag);

            for (var n = 0; n < count; ++n)
            {
                User32.GetMenuItemInfo(menu.Handle, n, true, ref itemInfo);

                if (itemInfo.hSubMenu == IntPtr.Zero)
                {
                    // If the item has no submenu we can't get the tag, so 
                    // check its ID to determine if it was added by the shell.
                    if (itemInfo.wID >= MCmdFirst) remove.Add(n);
                }
                else
                {
                    User32.GetMenuInfo(itemInfo.hSubMenu, ref menuInfo);
                    if (menuInfo.dwMenuData != tag) remove.Add(n);
                }
            }

            // Remove the unmanaged menu items.
            remove.Reverse();
            foreach (var position in remove)
            {
                User32.DeleteMenu(menu.Handle, position, MF.MF_BYPOSITION);
            }
        }

        private class MessageWindow : Control
        {
            private readonly ShellContextMenu _mParent;

            public MessageWindow(ShellContextMenu parent)
            {
                _mParent = parent;
            }

            protected override void WndProc(ref Message m)
            {
                if (!_mParent.HandleMenuMessage(ref m))
                {
                    base.WndProc(ref m);
                }
            }
        }
    }
}