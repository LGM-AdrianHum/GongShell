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

namespace GongSolutions.Shell
{
    /// <summary>
    ///     Holds a <see cref="ShellView" />'s navigation history.
    /// </summary>
    public class ShellHistory
    {
        private readonly List<ShellItem> _mHistory;
        private int _mCurrent;

        internal ShellHistory()
        {
            _mHistory = new List<ShellItem>();
        }

        /// <summary>
        ///     Gets the list of folders in the <see cref="ShellView" />'s
        ///     <b>Back</b> history.
        /// </summary>
        public ShellItem[] HistoryBack => _mHistory.GetRange(0, _mCurrent).ToArray();

        /// <summary>
        ///     Gets the list of folders in the <see cref="ShellView" />'s
        ///     <b>Forward</b> history.
        /// </summary>
        public ShellItem[] HistoryForward
        {
            get
            {
                if (CanNavigateForward)
                {
                    return _mHistory.GetRange(_mCurrent + 1,
                        _mHistory.Count - (_mCurrent + 1)).ToArray();
                }
                return new ShellItem[0];
            }
        }

        internal bool CanNavigateBack => _mCurrent > 0;

        internal bool CanNavigateForward => _mCurrent < _mHistory.Count - 1;

        internal ShellItem Current => _mHistory[_mCurrent];

        /// <summary>
        ///     Clears the shell history.
        /// </summary>
        public void Clear()
        {
            ShellItem current = null;

            if (_mHistory.Count > 0)
            {
                current = Current;
            }

            _mHistory.Clear();

            if (current != null)
            {
                Add(current);
            }
        }

        internal void Add(ShellItem folder)
        {
            while (_mCurrent < _mHistory.Count - 1)
            {
                _mHistory.RemoveAt(_mCurrent + 1);
            }

            _mHistory.Add(folder);
            _mCurrent = _mHistory.Count - 1;
        }

        internal ShellItem MoveBack()
        {
            if (_mCurrent == 0)
            {
                throw new InvalidOperationException("Cannot navigate back");
            }
            return _mHistory[--_mCurrent];
        }

        internal void MoveBack(ShellItem folder)
        {
            var index = _mHistory.IndexOf(folder);

            if ((index == -1) || (index >= _mCurrent))
            {
                throw new Exception(
                    "The requested folder could not be located in the " +
                    "'back' shell history");
            }

            _mCurrent = index;
        }

        internal ShellItem MoveForward()
        {
            if (_mCurrent == _mHistory.Count - 1)
            {
                throw new InvalidOperationException("Cannot navigate forward");
            }
            return _mHistory[++_mCurrent];
        }

        internal void MoveForward(ShellItem folder)
        {
            var index = _mHistory.IndexOf(folder, _mCurrent + 1);

            if (index == -1)
            {
                throw new Exception(
                    "The requested folder could not be located in the " +
                    "'forward' shell history");
            }

            _mCurrent = index;
        }
    }
}