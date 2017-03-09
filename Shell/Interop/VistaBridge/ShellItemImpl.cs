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
using System.Runtime.InteropServices;
using System.Text;

namespace GongSolutions.Shell.Interop.VistaBridge
{
    internal class ShellItemImpl : IDisposable, IShellItem
    {
        public ShellItemImpl(IntPtr pidl, bool owner)
        {
            if (owner)
            {
                Pidl = pidl;
            }
            else
            {
                Pidl = Shell32.ILClone(pidl);
            }
        }

        public IntPtr Pidl { get; }

        public void Dispose()
        {
            Dispose(true);
        }

        public IntPtr BindToHandler(IntPtr pbc, Guid bhid, Guid riid)
        {
            if (riid == typeof(IShellFolder).GUID)
            {
                return Marshal.GetIUnknownForObject(GetIShellFolder());
            }
            throw new InvalidCastException();
        }

        public HResult GetParent(out IShellItem ppsi)
        {
            var pidl = Shell32.ILClone(Pidl);
            if (Shell32.ILRemoveLastID(pidl))
            {
                ppsi = new ShellItemImpl(pidl, true);
                return HResult.S_OK;
            }
            ppsi = null;
            return HResult.MK_E_NOOBJECT;
        }

        public IntPtr GetDisplayName(SIGDN sigdnName)
        {
            if (sigdnName == SIGDN.FILESYSPATH)
            {
                var result = new StringBuilder(512);
                if (!Shell32.SHGetPathFromIDList(Pidl, result))
                    throw new ArgumentException();
                return Marshal.StringToHGlobalUni(result.ToString());
            }
            var parentFolder = GetParent().GetIShellFolder();
            var childPidl = Shell32.ILFindLastID(Pidl);
            var builder = new StringBuilder(512);
            var strret = new STRRET();

            parentFolder.GetDisplayNameOf(childPidl,
                (SHGNO) ((int) sigdnName & 0xffff), out strret);
            ShlWapi.StrRetToBuf(ref strret, childPidl, builder,
                (uint) builder.Capacity);
            return Marshal.StringToHGlobalUni(builder.ToString());
        }

        public SFGAO GetAttributes(SFGAO sfgaoMask)
        {
            var parentFolder = GetParent().GetIShellFolder();
            var result = sfgaoMask;

            parentFolder.GetAttributesOf(1,
                new[] {Shell32.ILFindLastID(Pidl)},
                ref result);
            return result & sfgaoMask;
        }

        public int Compare(IShellItem psi, SICHINT hint)
        {
            var other = (ShellItemImpl) psi;
            var myParent = GetParent();
            var theirParent = other.GetParent();

            if (Shell32.ILIsEqual(myParent.Pidl, theirParent.Pidl))
            {
                return myParent.GetIShellFolder().CompareIDs((SHCIDS) hint,
                    Shell32.ILFindLastID(Pidl),
                    Shell32.ILFindLastID(other.Pidl));
            }
            return 1;
        }

        ~ShellItemImpl()
        {
            Dispose(false);
        }

        protected void Dispose(bool dispose)
        {
            Shell32.ILFree(Pidl);
        }

        private ShellItemImpl GetParent()
        {
            var pidl = Shell32.ILClone(Pidl);

            if (Shell32.ILRemoveLastID(pidl))
            {
                return new ShellItemImpl(pidl, true);
            }
            return this;
        }

        private IShellFolder GetIShellFolder()
        {
            var desktop = Shell32.SHGetDesktopFolder();
            IntPtr desktopPidl;

            Shell32.SHGetSpecialFolderLocation(IntPtr.Zero, CSIDL.DESKTOP,
                out desktopPidl);
            ;

            if (Shell32.ILIsEqual(Pidl, desktopPidl))
            {
                return desktop;
            }
            IntPtr result;
            desktop.BindToObject(Pidl, IntPtr.Zero,
                typeof(IShellFolder).GUID, out result);
            return (IShellFolder) Marshal.GetTypedObjectForIUnknown(result,
                typeof(IShellFolder));
        }
    }
}