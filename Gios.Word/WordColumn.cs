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
	/// a Column of the WordTable
	/// </summary>
	public class WordColumn : WordCellRange
	{
		internal int Width,index;
		/// <summary>
		/// Gets the index of the column.
		/// </summary>
		public int Index
		{
			get
			{
				return this.index;
			}
		}
		private int compensatedWidth=0;
		internal WordColumn(WordTable WordTable,int index)
		{
			this.WordTable=WordTable;
			this.index=index;
			this.startColumn=index;
			this.endColumn=index;
			this.startRow=0;
			this.endRow=this.WordTable.rows-1;
		}
		/// <summary>
		/// sets the Relative Width of the Column. For example: if the relative widths of a 3 columns
		/// table are 10,10,30; the columns will respectivelly sized as 20%,20%,60% of the table size.
		/// </summary>
		/// <param name="RelativeWidth"></param>
		public void SetWidth(int RelativeWidth)
		{
			this.Width=RelativeWidth;
		}
		internal int CompensatedWidth
		{
			get
			{
				if (compensatedWidth!=0) return compensatedWidth;
				float sum=0;
				foreach (WordColumn pc in this.WordTable.rtfColumns)
				{
					sum+=pc.Width;
				}
				this.compensatedWidth=(int)(this.WordTable.width/sum)*this.Width;
				return this.compensatedWidth;
			}
		}
		
	}
}
