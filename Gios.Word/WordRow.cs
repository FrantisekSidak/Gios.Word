//============================================================================
//Gios Word.NET - A library for exporting Word Documents in C#
//using the Microsoft Rich Text Format (RTF) Specification
//Copyright (C) 2005  Paolo Gios - www.paologios.com
//
//This library is free software; you can redistribute it and/or
//modify it under the terms of the GNU Lesser General Public
//License as published by the Free Software Foundation; either
//version 2.1 of the License, or (at your option) any later version.
//
//This library is distributed in the hope that it will be useful,
//but WITHOUT ANY WARRANTY; without even the implied warranty of
//MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
//Lesser General Public License for more details.
//
//You should have received a copy of the GNU Lesser General Public
//License along with this library; if not, write to the Free Software
//Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
//=============================================================================
using System;
using System.Collections;
using System.Drawing;

namespace Gios.Word
{
	/// <summary>
	/// row of a WordTable
	/// </summary>
	public class WordRow : WordCellRange
	{
		internal int height=0,index;
		/// <summary>
		/// Gets the index of the row.
		/// </summary>
		public int Index
		{
			get
			{
				return this.index;
			}
		}
		/// <summary>
		/// Sets the row height in twips.
		/// </summary>
		/// <param name="Height"></param>
		public void SetRowHeight(int Height)
		{
			this.height=Height;
		}
		internal WordRow(WordTable WordTable,int index)
		{
			this.WordTable=WordTable;
			this.index=index;
			this.startColumn=0;
			this.endColumn=this.WordTable.columns-1;
			this.startRow=index;
			this.endRow=index;
		}
		/// <summary>
		/// Gets a cell specifying the column.
		/// </summary>
		public WordCell this[int column]
		{
			get
			{
				return this.WordTable.Cell(this.index,column);
			}
		}
		
	}
}
