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
using System.Runtime.InteropServices;
using System.Text;
using GongSolutions.Shell.Interop;

namespace GongSolutions.Shell
{
    /// <summary>
    /// </summary>
    public class KnownFolderManager : IEnumerable<KnownFolder>
    {
        private readonly IKnownFolderManager _mComInterface;
        private readonly Dictionary<string, KnownFolder> _mNameIndex;
        private readonly Dictionary<string, KnownFolder> _mPathIndex;

        /// <summary>
        /// </summary>
        public KnownFolderManager()
        {
            if (Environment.OSVersion.Version.Major >= 6)
            {
                // ReSharper disable once SuspiciousTypeConversion.Global
                _mComInterface = (IKnownFolderManager) new CoClass.KnownFolderManager();
            }
            else
            {
                _mNameIndex = new Dictionary<string, KnownFolder>();
                _mPathIndex = new Dictionary<string, KnownFolder>();

                AddFolder("Common Desktop", CSIDL.COMMON_DESKTOPDIRECTORY);
                AddFolder("Desktop", CSIDL.DESKTOP);
                AddFolder("Personal", CSIDL.PERSONAL);
                AddFolder("Recent", CSIDL.RECENT);
                AddFolder("MyComputerFolder", CSIDL.DRIVES);
                AddFolder("My Pictures", CSIDL.MYPICTURES);
                AddFolder("ProgramFilesCommon", CSIDL.PROGRAM_FILES_COMMON);
                AddFolder("Windows", CSIDL.WINDOWS);
            }
        }

        public IEnumerator<KnownFolder> GetEnumerator()
        {
            if (_mComInterface != null)
            {
                IntPtr buffer;
                uint count;
                _mComInterface.GetFolderIds(out buffer, out count);

                KnownFolder[] results;
                try
                {
                    results = new KnownFolder[count];
                    var p = buffer;

                    for (uint n = 0; n < count; ++n)
                    {
                        var guid = (Guid) Marshal.PtrToStructure(p, typeof(Guid));
                        results[n] = GetFolder(guid);
                        p = (IntPtr) ((int) p + Marshal.SizeOf(typeof(Guid)));
                    }
                }
                finally
                {
                    Marshal.FreeCoTaskMem(buffer);
                }

                foreach (var f in results)
                {
                    yield return f;
                }
            }
            else
            {
                foreach (var f in _mNameIndex.Values)
                {
                    yield return f;
                }
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return ((IEnumerable<KnownFolder>) this).GetEnumerator();
        }

        /// <summary>
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        public KnownFolder FindNearestParent(ShellItem item)
        {
            if (_mComInterface != null)
            {
                IKnownFolder iKnownFolder;

                if (item.IsFileSystem)
                {
                    if (_mComInterface.FindFolderFromPath(item.FileSystemPath,
                        FFFP_MODE.NEARESTPARENTMATCH, out iKnownFolder)
                        == HResult.S_OK)
                    {
                        return CreateFolder(iKnownFolder);
                    }
                }
                else
                {
                    if (_mComInterface.FindFolderFromIDList(item.Pidl, out iKnownFolder)
                        == HResult.S_OK)
                    {
                        return CreateFolder(iKnownFolder);
                    }
                }
            }
            else
            {
                if (item.IsFileSystem)
                {
                    foreach (var i in _mPathIndex)
                    {
                        if ((i.Key != string.Empty) &&
                            item.FileSystemPath.StartsWith(i.Key))
                        {
                            return i.Value;
                        }
                    }
                }
                else
                {
                    foreach (var i in _mNameIndex)
                    {
                        if (item == i.Value.CreateShellItem())
                        {
                            return i.Value;
                        }
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// </summary>
        /// <param name="guid"></param>
        /// <returns></returns>
        public KnownFolder GetFolder(Guid guid)
        {
            return CreateFolder(_mComInterface.GetFolder(guid));
        }

        /// <summary>
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        /// <exception cref="InvalidOperationException"></exception>
        public KnownFolder GetFolder(string name)
        {
            if (_mComInterface != null)
            {
                IKnownFolder iKnownFolder;

                if (_mComInterface.GetFolderByName(name, out iKnownFolder)
                    == HResult.S_OK)
                {
                    return CreateFolder(iKnownFolder);
                }
            }
            else
            {
                return _mNameIndex[name];
            }

            throw new InvalidOperationException("Unknown shell folder: " + name);
        }

        private void AddFolder(string name, CSIDL csidl)
        {
            var folder = CreateFolder(csidl, name);

            _mNameIndex.Add(folder.Name, folder);

            if (folder.ParsingName != string.Empty)
            {
                _mPathIndex.Add(folder.ParsingName, folder);
            }
        }

        private static KnownFolder CreateFolder(CSIDL csidl, string name)
        {
            var path = new StringBuilder(512);

            if (Shell32.SHGetFolderPath(IntPtr.Zero, csidl, IntPtr.Zero, 0, path) == HResult.S_OK)
            {
                return new KnownFolder(csidl, name, path.ToString());
            }
            return new KnownFolder(csidl, name, string.Empty);
        }

        private static KnownFolder CreateFolder(IKnownFolder iface)
        {
            var def = iface.GetFolderDefinition();

            try
            {
                return new KnownFolder(iface,
                    Marshal.PtrToStringUni(def.pszName),
                    Marshal.PtrToStringUni(def.pszParsingName));
            }
            finally
            {
                Marshal.FreeCoTaskMem(def.pszName);
                Marshal.FreeCoTaskMem(def.pszDescription);
                Marshal.FreeCoTaskMem(def.pszRelativePath);
                Marshal.FreeCoTaskMem(def.pszParsingName);
                Marshal.FreeCoTaskMem(def.pszTooltip);
                Marshal.FreeCoTaskMem(def.pszLocalizedName);
                Marshal.FreeCoTaskMem(def.pszIcon);
                Marshal.FreeCoTaskMem(def.pszSecurity);
            }
        }

        // ReSharper disable once UnusedMember.Local
        private struct PathIndexEntry
        {
            public PathIndexEntry(string name, CSIDL csidl)
            {
                Name = name;
                Csidl = csidl;
            }

            // ReSharper disable once UnusedAutoPropertyAccessor.Local
            // ReSharper disable once MemberCanBePrivate.Local
            public string Name { get; set; }

            // ReSharper disable once MemberCanBePrivate.Local
            // ReSharper disable once UnusedAutoPropertyAccessor.Local
            public CSIDL Csidl { get; set; }
        }
    }

    /// <summary>
    /// </summary>
    public class KnownFolder
    {
        private readonly IKnownFolder _mComInterface;
        private readonly CSIDL _mCsidl;

        /// <summary>
        /// </summary>
        /// <param name="iface"></param>
        /// <param name="name"></param>
        /// <param name="parsingName"></param>
        public KnownFolder(IKnownFolder iface, string name, string parsingName)
        {
            _mComInterface = iface;
            Name = name;
            ParsingName = parsingName;
        }

        /// <summary>
        /// </summary>
        /// <param name="csidl"></param>
        /// <param name="name"></param>
        /// <param name="parsingName"></param>
        public KnownFolder(CSIDL csidl, string name, string parsingName)
        {
            _mCsidl = csidl;
            Name = name;
            ParsingName = parsingName;
        }

        /// <summary>
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// </summary>
        public string ParsingName { get; set; }

        /// <summary>
        /// </summary>
        /// <returns></returns>
        public ShellItem CreateShellItem()
        {
            if (_mComInterface != null)
            {
                return new ShellItem(_mComInterface.GetShellItem(0,
                    typeof(IShellItem).GUID));
            }
            return new ShellItem((Environment.SpecialFolder) _mCsidl);
        }
    }
}