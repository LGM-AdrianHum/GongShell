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
    internal class FilterItem
    {
        public string Caption;
        public string Filter;

        public FilterItem(string caption, string filter)
        {
            Caption = caption;
            Filter = filter;
        }

        public bool Contains(string filter)
        {
            var filters = Filter.Split(',');

            foreach (var s in filters)
            {
                if (filter == s.Trim()) return true;
            }

            return false;
        }

        public override string ToString()
        {
            var filterString = $" ({Filter})";

            if (Caption.EndsWith(filterString))
            {
                return Caption;
            }
            return Caption + filterString;
        }

        public static FilterItem[] ParseFilterString(string filterString)
        {
            int dummy;
            return ParseFilterString(filterString, string.Empty, out dummy);
        }

        public static FilterItem[] ParseFilterString(string filterString,
            string existing,
            out int existingIndex)
        {
            var result = new List<FilterItem>();

            existingIndex = -1;

            var items = filterString != string.Empty ? filterString.Split('|') : new string[0];

            if (items.Length%2 != 0)
            {
                throw new ArgumentException(
                    "Filter string you provided is not valid. The filter " +
                    "string must contain a description of the filter, " +
                    "followed by the vertical bar (|) and the filter pattern." +
                    "The strings for different filtering options must also be " +
                    "separated by the vertical bar. Example: " +
                    "\"Text files|*.txt|All files|*.*\"");
            }

            for (var n = 0; n < items.Length; n += 2)
            {
                var item = new FilterItem(items[n], items[n + 1]);
                result.Add(item);
                if (item.Filter == existing) existingIndex = result.Count - 1;
            }

            return result.ToArray();
        }
    }
}