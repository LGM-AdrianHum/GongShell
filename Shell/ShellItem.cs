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
using System.ComponentModel.Design.Serialization;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using GongSolutions.Shell.Interop;
using GongSolutions.Shell.Interop.VistaBridge;
using IDataObject = System.Runtime.InteropServices.ComTypes.IDataObject;
using IDropTarget = GongSolutions.Shell.Interop.IDropTarget;

namespace GongSolutions.Shell
{
    /// <summary>
    ///     Represents an item in the Windows Shell namespace.
    /// </summary>
    [TypeConverter(typeof(ShellItemConverter))]
#pragma warning disable 660,661
    public class ShellItem : IEnumerable<ShellItem>
#pragma warning restore 660,661
    {
        private static ShellItem _mDesktop;

        /// <summary>
        ///     Initializes a new instance of the <see cref="ShellItem" /> class.
        /// </summary>
        /// <remarks>
        ///     Takes a <see cref="Uri" /> containing the location of the ShellItem.
        ///     This constructor accepts URIs using two schemes:
        ///     - file: A file or folder in the computer's filesystem, e.g.
        ///     file:///D:/Folder
        ///     - shell: A virtual folder, or a file or folder referenced from
        ///     a virtual folder, e.g. shell:///Personal/file.txt
        /// </remarks>
        /// <param name="uri">
        ///     A <see cref="Uri" /> containing the location of the ShellItem.
        /// </param>
        public ShellItem(Uri uri)
        {
            Initialize(uri);
        }

        /// <summary>
        ///     Initializes a new instance of the <see cref="ShellItem" /> class.
        /// </summary>
        /// <remarks>
        ///     Takes a <see cref="string" /> containing the location of the ShellItem.
        ///     This constructor accepts URIs using two schemes:
        ///     - file: A file or folder in the computer's filesystem, e.g.
        ///     file:///D:/Folder
        ///     - shell: A virtual folder, or a file or folder referenced from
        ///     a virtual folder, e.g. shell:///Personal/file.txt
        /// </remarks>
        /// <param name="path">
        ///     A string containing a Uri with the location of the ShellItem.
        /// </param>
        public ShellItem(string path)
        {
            Initialize(new Uri(path));
        }

        /// <summary>
        ///     Initializes a new instance of the <see cref="ShellItem" /> class.
        /// </summary>
        /// <remarks>
        ///     Takes an <see cref="Environment.SpecialFolder" /> containing the
        ///     location of the folder.
        /// </remarks>
        /// <param name="folder">
        ///     An <see cref="Environment.SpecialFolder" /> containing the
        ///     location of the folder.
        /// </param>
        public ShellItem(Environment.SpecialFolder folder)
        {
            IntPtr pidl;

            if (Shell32.SHGetSpecialFolderLocation(IntPtr.Zero,
                (CSIDL) folder, out pidl) == HResult.S_OK)
            {
                try
                {
                    ComInterface = CreateItemFromIdList(pidl);
                }
                finally
                {
                    Shell32.ILFree(pidl);
                }
            }
            else
            {
                // SHGetSpecialFolderLocation does not support many common
                // CSIDL values on Windows 98, but SHGetFolderPath in 
                // ShFolder.dll does, so fall back to it if necessary. We
                // try SHGetSpecialFolderLocation first because it returns
                // a PIDL which is preferable to a path as it can express
                // virtual folder locations.
                var path = new StringBuilder();
                Marshal.ThrowExceptionForHR((int) Shell32.SHGetFolderPath(
                    IntPtr.Zero, (CSIDL) folder, IntPtr.Zero, 0, path));
                ComInterface = CreateItemFromParsingName(path.ToString());
            }
        }

        /// <summary>
        ///     Initializes a new instance of the <see cref="ShellItem" /> class.
        /// </summary>
        /// <remarks>
        ///     Creates a ShellItem which is a named child of <paramref name="parent" />.
        /// </remarks>
        /// <param name="parent">
        ///     The parent folder of the item.
        /// </param>
        /// <param name="name">
        ///     The name of the child item.
        /// </param>
        public ShellItem(ShellItem parent, string name)
        {
            if (parent.IsFileSystem)
            {
                // If the parent folder is in the file system, our best 
                // chance of success is to use the FileSystemPath to 
                // create the new item. Folders other than Desktop don't 
                // seem to implement ParseDisplayName properly.
                ComInterface = CreateItemFromParsingName(
                    Path.Combine(parent.FileSystemPath, name));
            }
            else
            {
                var folder = parent.GetIShellFolder();
                uint eaten;
                IntPtr pidl;
                uint attributes = 0;

                folder.ParseDisplayName(IntPtr.Zero, IntPtr.Zero,
                    name, out eaten, out pidl, ref attributes);

                try
                {
                    ComInterface = CreateItemFromIdList(pidl);
                }
                finally
                {
                    Shell32.ILFree(pidl);
                }
            }
        }

        internal ShellItem(IntPtr pidl)
        {
            ComInterface = CreateItemFromIdList(pidl);
        }

        internal ShellItem(ShellItem parent, IntPtr pidl)
        {
            ComInterface = CreateItemWithParent(parent, pidl);
        }


        /// <summary>
        ///     Initializes a new instance of the <see cref="ShellItem" /> class.
        /// </summary>
        /// <param name="comInterface">
        ///     An <see cref="IShellItem" /> representing the folder.
        /// </param>
        public ShellItem(IShellItem comInterface)
        {
            ComInterface = comInterface;
        }

        /// <summary>
        ///     Gets a child item.
        /// </summary>
        /// <param name="name">
        ///     The name of the child item.
        /// </param>
        public ShellItem this[string name] => new ShellItem(this, name);

        /// <summary>
        ///     Gets the underlying <see cref="IShellItem" /> COM interface.
        /// </summary>
        public IShellItem ComInterface { get; set; }

        /// <summary>
        ///     Gets the normal display name of the item.
        /// </summary>
        public string DisplayName
        {
            get { return GetDisplayName(SIGDN.NORMALDISPLAY); }
        }

        /// <summary>
        ///     Gets the file system path of the item.
        /// </summary>
        public string FileSystemPath
        {
            get { return GetDisplayName(SIGDN.FILESYSPATH); }
        }

        /// <summary>
        ///     Gets a value indicating whether the item has subfolders.
        /// </summary>
        public bool HasSubFolders
        {
            get { return ComInterface.GetAttributes(SFGAO.HASSUBFOLDER) != 0; }
        }

        /// <summary>
        ///     Gets a value indicating whether the item is a file system item.
        /// </summary>
        public bool IsFileSystem
        {
            get { return ComInterface.GetAttributes(SFGAO.FILESYSTEM) != 0; }
        }

        /// <summary>
        ///     Gets a value indicating whether the item is a file system item
        ///     or the child of a file system item.
        /// </summary>
        public bool IsFileSystemAncestor
        {
            get { return ComInterface.GetAttributes(SFGAO.FILESYSANCESTOR) != 0; }
        }

        /// <summary>
        ///     Gets a value indicating whether the item is a folder.
        /// </summary>
        public bool IsFolder
        {
            get { return ComInterface.GetAttributes(SFGAO.FOLDER) != 0; }
        }

        /// <summary>
        ///     Gets a value indicating whether the item is read-only.
        /// </summary>
        public bool IsReadOnly
        {
            get { return ComInterface.GetAttributes(SFGAO.READONLY) != 0; }
        }

        /// <summary>
        ///     Gets the item's parent.
        /// </summary>
        public ShellItem Parent
        {
            get
            {
                IShellItem item;
                var result = ComInterface.GetParent(out item);

                if (result == HResult.S_OK)
                {
                    return new ShellItem(item);
                }
                if (result == HResult.MK_E_NOOBJECT)
                {
                    return null;
                }
                Marshal.ThrowExceptionForHR((int) result);
                return null;
            }
        }

        /// <summary>
        ///     Gets the item's parsing name.
        /// </summary>
        public string ParsingName
        {
            get { return GetDisplayName(SIGDN.DESKTOPABSOLUTEPARSING); }
        }

        /// <summary>
        ///     Gets a PIDL representing the item.
        /// </summary>
        public IntPtr Pidl
        {
            get { return GetIdListFromObject(ComInterface); }
        }

        /// <summary>
        ///     Gets the item's shell icon.
        /// </summary>
        public Icon ShellIcon
        {
            get
            {
                var info = new SHFILEINFO();
                var result = Shell32.SHGetFileInfo(Pidl, 0, out info,
                    Marshal.SizeOf(info),
                    SHGFI.ADDOVERLAYS | SHGFI.ICON |
                    SHGFI.SHELLICONSIZE | SHGFI.PIDL);

                if (result == IntPtr.Zero)
                {
                    throw new Exception("Error retreiving shell folder icon");
                }

                return Icon.FromHandle(info.hIcon);
            }
        }

        /// <summary>
        ///     Gets the item's tooltip text.
        /// </summary>
        public string ToolTipText
        {
            get
            {
                IntPtr result;
                IQueryInfo queryInfo;
                IntPtr infoTipPtr;
                string infoTip;

                try
                {
                    var relativePidl = Shell32.ILFindLastID(Pidl);
                    Parent.GetIShellFolder().GetUIObjectOf(IntPtr.Zero, 1,
                        new[] {relativePidl},
                        typeof(IQueryInfo).GUID, 0, out result);
                }
                catch (Exception)
                {
                    return string.Empty;
                }

                queryInfo = (IQueryInfo)
                    Marshal.GetTypedObjectForIUnknown(result,
                        typeof(IQueryInfo));
                queryInfo.GetInfoTip(0, out infoTipPtr);
                infoTip = Marshal.PtrToStringUni(infoTipPtr);
                Ole32.CoTaskMemFree(infoTipPtr);
                return infoTip;
            }
        }

        /// <summary>
        ///     Gets the Desktop folder.
        /// </summary>
        public static ShellItem Desktop
        {
            get
            {
                if (_mDesktop == null)
                {
                    IShellItem item;
                    IntPtr pidl;

                    Shell32.SHGetSpecialFolderLocation(
                        IntPtr.Zero, (CSIDL) Environment.SpecialFolder.Desktop,
                        out pidl);

                    try
                    {
                        item = CreateItemFromIdList(pidl);
                    }
                    finally
                    {
                        Shell32.ILFree(pidl);
                    }

                    _mDesktop = new ShellItem(item);
                }
                return _mDesktop;
            }
        }

        internal static bool RunningVista
        {
            get { return Environment.OSVersion.Version.Major >= 6; }
        }

        /// <summary>
        ///     Returns an enumerator detailing the child items of the
        ///     <see cref="ShellItem" />.
        /// </summary>
        /// <remarks>
        ///     This method returns all child item including hidden
        ///     items.
        /// </remarks>
        /// <returns>
        ///     An enumerator over all child items.
        /// </returns>
        public IEnumerator<ShellItem> GetEnumerator()
        {
            return GetEnumerator(SHCONTF.FOLDERS | SHCONTF.INCLUDEHIDDEN |
                                 SHCONTF.NONFOLDERS);
        }

        /// <summary>
        ///     Returns an enumerator detailing the child items of the
        ///     <see cref="ShellItem" />.
        /// </summary>
        /// <remarks>
        ///     This method returns all child item including hidden
        ///     items.
        /// </remarks>
        /// <returns>
        ///     An enumerator over all child items.
        /// </returns>
        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        /// <summary>
        /// </summary>
        /// <param name="other"></param>
        /// <returns></returns>
        protected bool Equals(ShellItem other)
        {
            return Equals(ComInterface, other.ComInterface);
        }

        /// <summary>
        ///     Compares two <see cref="IShellItem" />s. The comparison is carried
        ///     out by display order.
        /// </summary>
        /// <param name="item">
        ///     The item to compare.
        /// </param>
        /// <returns>
        ///     0 if the two items are equal. A negative number if
        ///     <see langword="this" /> is before <paramref name="item" /> in
        ///     display order. A positive number if
        ///     <see langword="this" /> comes after <paramref name="item" /> in
        ///     display order.
        /// </returns>
        public int Compare(ShellItem item)
        {
            var result = ComInterface.Compare(item.ComInterface,
                SICHINT.DISPLAY);
            return result;
        }

        /// <summary>
        ///     Determines whether two <see cref="ShellItem" />s refer to
        ///     the same shell folder.
        /// </summary>
        /// <param name="obj">
        ///     The item to compare.
        /// </param>
        /// <returns>
        ///     <see langword="true" /> if the two objects refer to the same
        ///     folder, <see langword="false" /> otherwise.
        /// </returns>
        public override bool Equals(object obj)
        {
            if (obj is ShellItem)
            {
                var otherItem = (ShellItem) obj;
                var result = ComInterface.Compare(otherItem.ComInterface,
                    SICHINT.DISPLAY) == 0;

                // Sometimes, folders are reported as being unequal even when
                // they refer to the same folder, so double check by comparing
                // the file system paths. (This was showing up on Windows XP in 
                // the SpecialFolders() test)
                if (!result)
                {
                    result = IsFileSystem && otherItem.IsFileSystem &&
                             (FileSystemPath == otherItem.FileSystemPath);
                }

                return result;
            }
            return false;
        }

        /// <summary>
        ///     Returns the name of the item in the specified style.
        /// </summary>
        /// <param name="sigdn">
        ///     The style of display name to return.
        /// </param>
        /// <returns>
        ///     A string containing the display name of the item.
        /// </returns>
        public string GetDisplayName(SIGDN sigdn)
        {
            var resultPtr = ComInterface.GetDisplayName(sigdn);
            var result = Marshal.PtrToStringUni(resultPtr);
            Marshal.FreeCoTaskMem(resultPtr);
            return result;
        }

        /// <summary>
        ///     Returns an enumerator detailing the child items of the
        ///     <see cref="ShellItem" />.
        /// </summary>
        /// <param name="filter">
        ///     A filter describing the types of child items to be included.
        /// </param>
        /// <returns>
        ///     An enumerator over all child items.
        /// </returns>
        public IEnumerator<ShellItem> GetEnumerator(SHCONTF filter)
        {
            var folder = GetIShellFolder();
            var enumId = GetIEnumIdList(folder, filter);
            uint count;
            IntPtr pidl;

            if (enumId == null)
            {
                yield break;
            }

            var result = enumId.Next(1, out pidl, out count);
            while (result == HResult.S_OK)
            {
                yield return new ShellItem(this, pidl);
                Shell32.ILFree(pidl);
                result = enumId.Next(1, out pidl, out count);
            }

            if (result != HResult.S_FALSE)
            {
                Marshal.ThrowExceptionForHR((int) result);
            }
        }

        /// <summary>
        ///     Returns an <see cref="System.Runtime.InteropServices.ComTypes.IDataObject" /> representing the
        ///     item. This object is used in drag and drop operations.
        /// </summary>
        public IDataObject GetIDataObject()
        {
            var result = ComInterface.BindToHandler(IntPtr.Zero,
                BHID.SFUIObject, typeof(IDataObject).GUID);
            return (IDataObject) Marshal.GetTypedObjectForIUnknown(result,
                typeof(IDataObject));
        }

        /// <summary>
        ///     Returns an <see cref="IDropTarget" /> representing the
        ///     item. This object is used in drag and drop operations.
        /// </summary>
        public IDropTarget GetIDropTarget(Control control)
        {
            var result = GetIShellFolder().CreateViewObject(control.Handle,
                typeof(IDropTarget).GUID);
            return (IDropTarget) Marshal.GetTypedObjectForIUnknown(result,
                typeof(IDropTarget));
        }

        /// <summary>
        ///     Returns an <see cref="IShellFolder" /> representing the
        ///     item.
        /// </summary>
        public IShellFolder GetIShellFolder()
        {
            var result = ComInterface.BindToHandler(IntPtr.Zero,
                BHID.SFObject, typeof(IShellFolder).GUID);
            return (IShellFolder) Marshal.GetTypedObjectForIUnknown(result,
                typeof(IShellFolder));
        }

        /// <summary>
        ///     Gets the index in the system image list of the icon representing
        ///     the item.
        /// </summary>
        /// <param name="type">
        ///     The type of icon to retrieve.
        /// </param>
        /// <param name="flags">
        ///     Flags detailing additional information to be conveyed by the icon.
        /// </param>
        /// <returns></returns>
        public int GetSystemImageListIndex(ShellIconType type, ShellIconFlags flags)
        {
            var info = new SHFILEINFO();
            var result = Shell32.SHGetFileInfo(Pidl, 0, out info,
                Marshal.SizeOf(info),
                SHGFI.ICON | SHGFI.SYSICONINDEX | SHGFI.OVERLAYINDEX | SHGFI.PIDL |
                (SHGFI) type | (SHGFI) flags);

            if (result == IntPtr.Zero)
            {
                throw new Exception("Error retreiving shell folder icon");
            }

            return info.iIcon;
        }

        /// <summary>
        ///     Tests whether the <see cref="ShellItem" /> is the immediate parent
        ///     of another item.
        /// </summary>
        /// <param name="item">
        ///     The potential child item.
        /// </param>
        public bool IsImmediateParentOf(ShellItem item)
        {
            return IsFolder && Shell32.ILIsParent(Pidl, item.Pidl, true);
        }

        /// <summary>
        ///     Tests whether the <see cref="ShellItem" /> is the parent of
        ///     another item.
        /// </summary>
        /// <param name="item">
        ///     The potential child item.
        /// </param>
        public bool IsParentOf(ShellItem item)
        {
            return IsFolder && Shell32.ILIsParent(Pidl, item.Pidl, false);
        }

        /// <summary>
        ///     Returns a string representation of the <see cref="ShellItem" />.
        /// </summary>
        public override string ToString()
        {
            return ToUri().ToString();
        }

        /// <summary>
        ///     Returns a URI representation of the <see cref="ShellItem" />.
        /// </summary>
        public Uri ToUri()
        {
            var manager = new KnownFolderManager();
            var path = new StringBuilder("shell:///");
            var knownFolder = manager.FindNearestParent(this);

            if (knownFolder != null)
            {
                var folders = new List<string>();
                var knownFolderItem = knownFolder.CreateShellItem();
                var item = this;

                while (item != knownFolderItem)
                {
                    folders.Add(item.GetDisplayName(SIGDN.PARENTRELATIVEPARSING));
                    item = item.Parent;
                }

                folders.Reverse();
                path.Append(knownFolder.Name);
                foreach (var s in folders)
                {
                    path.Append('/');
                    path.Append(s);
                }

                return new Uri(path.ToString());
            }
            return new Uri(FileSystemPath);
        }

        /// <summary>
        ///     Tests if two <see cref="ShellItem" />s refer to the same folder.
        /// </summary>
        /// <param name="a">
        ///     The first folder.
        /// </param>
        /// <param name="b">
        ///     The second folder.
        /// </param>
        public static bool operator !=(ShellItem a, ShellItem b)
        {
            if (ReferenceEquals(a, null))
            {
                return !ReferenceEquals(b, null);
            }
            return !a.Equals(b);
        }

        /// <summary>
        ///     Tests if two <see cref="ShellItem" />s refer to the same folder.
        /// </summary>
        /// <param name="a">
        ///     The first folder.
        /// </param>
        /// <param name="b">
        ///     The second folder.
        /// </param>
        public static bool operator ==(ShellItem a, ShellItem b)
        {
            if (ReferenceEquals(a, null))
            {
                return ReferenceEquals(b, null);
            }
            return a.Equals(b);
        }

        private void Initialize(Uri uri)
        {
            switch (uri.Scheme)
            {
                case "file":
                    ComInterface = CreateItemFromParsingName(uri.LocalPath);
                    break;
                case "shell":
                    InitializeFromShellUri(uri);
                    break;
                default:
                    throw new InvalidOperationException("Invalid uri scheme");
            }
        }

        private void InitializeFromShellUri(Uri uri)
        {
            var manager = new KnownFolderManager();
            var path = uri.GetComponents(UriComponents.Path, UriFormat.Unescaped);
            string knownFolder;
            string restOfPath;
            var separatorIndex = path.IndexOf('/');

            if (separatorIndex != -1)
            {
                knownFolder = path.Substring(0, separatorIndex);
                restOfPath = path.Substring(separatorIndex + 1);
            }
            else
            {
                knownFolder = path;
                restOfPath = string.Empty;
            }

            ComInterface = manager.GetFolder(knownFolder).CreateShellItem().ComInterface;

            if (restOfPath != string.Empty)
            {
                ComInterface = this[restOfPath.Replace('/', '\\')].ComInterface;
            }
        }

        private static IShellItem CreateItemFromIdList(IntPtr pidl)
        {
            if (RunningVista)
            {
                return Shell32.SHCreateItemFromIDList(pidl,
                    typeof(IShellItem).GUID);
            }
            return new ShellItemImpl(
                pidl, false);
        }

        private static IShellItem CreateItemFromParsingName(string path)
        {
            if (RunningVista)
            {
                return Shell32.SHCreateItemFromParsingName(path, IntPtr.Zero,
                    typeof(IShellItem).GUID);
            }
            var desktop = Desktop.GetIShellFolder();
            uint attributes = 0;
            uint eaten;
            IntPtr pidl;

            desktop.ParseDisplayName(IntPtr.Zero, IntPtr.Zero,
                path, out eaten, out pidl, ref attributes);
            return new ShellItemImpl(
                pidl, true);
        }

        private static IShellItem CreateItemWithParent(ShellItem parent, IntPtr pidl)
        {
            if (RunningVista)
            {
                return Shell32.SHCreateItemWithParent(IntPtr.Zero,
                    parent.GetIShellFolder(), pidl, typeof(IShellItem).GUID);
            }
            var impl =
                (ShellItemImpl) parent.ComInterface;
            return new ShellItemImpl(
                Shell32.ILCombine(impl.Pidl, pidl), true);
        }

        private static IntPtr GetIdListFromObject(IShellItem item)
        {
            if (RunningVista)
            {
                return Shell32.SHGetIDListFromObject(item);
            }
            return ((ShellItemImpl) item).Pidl;
        }

        private static IEnumIDList GetIEnumIdList(IShellFolder folder, SHCONTF flags)
        {
            IEnumIDList result;

            if (folder.EnumObjects(IntPtr.Zero, flags, out result) == HResult.S_OK)
            {
                return result;
            }
            return null;
        }
    }

    internal class ShellItemConverter : TypeConverter
    {
        public override bool CanConvertFrom(ITypeDescriptorContext context,
            Type sourceType)
        {
            if (sourceType == typeof(string))
            {
                return true;
            }
            return base.CanConvertFrom(context, sourceType);
        }

        public override bool CanConvertTo(ITypeDescriptorContext context,
            Type destinationType)
        {
            if (destinationType == typeof(InstanceDescriptor))
            {
                return true;
            }
            return base.CanConvertTo(context, destinationType);
        }

        public override object ConvertFrom(ITypeDescriptorContext context,
            CultureInfo culture,
            object value)
        {
            if (value is string)
            {
                var s = (string) value;

                if (s.Length == 0)
                {
                    return ShellItem.Desktop;
                }
                return new ShellItem(s);
            }
            return base.ConvertFrom(context, culture, value);
        }

        public override object ConvertTo(ITypeDescriptorContext context,
            CultureInfo culture,
            object value,
            Type destinationType)
        {
            if (value is ShellItem)
            {
                var uri = ((ShellItem) value).ToUri();

                if (destinationType == typeof(string))
                {
                    if (uri.Scheme == "file")
                    {
                        return uri.LocalPath;
                    }
                    return uri.ToString();
                }
                if (destinationType == typeof(InstanceDescriptor))
                {
                    return new InstanceDescriptor(
                        typeof(ShellItem).GetConstructor(new[] {typeof(string)}),
                        new object[] {uri.ToString()});
                }
            }
            return base.ConvertTo(context, culture, value, destinationType);
        }
    }

    /// <summary>
    ///     Enumerates the types of shell icons.
    /// </summary>
    public enum ShellIconType
    {
        /// <summary>The system large icon type</summary>
        LargeIcon = SHGFI.LARGEICON,

        /// <summary>The system shell icon type</summary>
        ShellIcon = SHGFI.SHELLICONSIZE,

        /// <summary>The system small icon type</summary>
        SmallIcon = SHGFI.SMALLICON
    }

    /// <summary>
    ///     Enumerates the optional styles that can be applied to shell icons.
    /// </summary>
    [Flags]
    public enum ShellIconFlags
    {
        /// <summary>The icon is displayed opened.</summary>
        OpenIcon = SHGFI.OPENICON,

        /// <summary>Get the overlay for the icon as well.</summary>
        OverlayIndex = SHGFI.OVERLAYINDEX
    }
}