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
using System.Drawing;
using System.Collections;

namespace Gios.Word
{
	/// <summary>
	/// a range of cell of the WordTable.
	/// </summary>
	public class WordCellRange
	{
		internal int startRow,endRow,startColumn,endColumn;
		internal WordTable WordTable;
		internal WordCellRange(){}
		internal WordCellRange(WordTable WordTable,int startRow,int startColumn,int endRow,int endColumn)
		{
			this.WordTable=WordTable;
			this.startColumn=startColumn;
			this.startRow=startRow;
			this.endColumn=endColumn;
			this.endRow=endRow;
		}
		/// <summary>
		/// Sets this padding to each cell of the cellrange.
		/// </summary>
		/// <param name="CellPadding"></param>
		public void SetCellPadding(int CellPadding)
		{
			foreach (WordCell rc in this.Cells)	rc.SetCellPadding(CellPadding);
		}
		/// <summary>
		/// the entire collection of cells.
		/// </summary>
		public ArrayList Cells
		{
			get
			{
				ArrayList al=new ArrayList();
				for (int r=startRow;r<=endRow;r++)
					for (int c=startColumn;c<=endColumn;c++)
						al.Add(this.WordTable.Cell(r,c));
                return al;
			}
		}
		/// <summary>
		/// Sets this background color to each cell of the cellrange. 
		/// </summary>
		/// <param name="BackgroundColor"></param>
		public void SetBackgroundColor(Color BackgroundColor)
		{
			foreach (WordCell rc in this.Cells)	rc.BackgroundColor=BackgroundColor;
		}
		/// <summary>
		/// Sets those background colors to each cell of the cellrange.
		/// </summary>
		/// <param name="BackgroundColor"></param>
		/// <param name="AlternateBackgroundColor"></param>
		public void SetBackgroundColor(Color BackgroundColor,Color AlternateBackgroundColor)
		{
			this.SetBackgroundColor(AlternateBackgroundColor);
			for (int r=this.startRow;r<=this.endRow;r+=2)
				for (int c=this.startColumn;c<=this.endColumn;c++)
					this.WordTable.Cell(r,c).BackgroundColor=BackgroundColor;
		}
		/// <summary>
		/// Sets this foreground color to each cell of the cellrange.
		/// </summary>
		/// <param name="Color"></param>
		public void SetForegroundColor(Color Color)
		{
			foreach (WordCell rc in this.Cells)	rc.SetForegroundColor(Color);
		}
		/// <summary>
		/// Sets those colors to each cell of the cellrange.
		/// </summary>
		/// <param name="ForegroundColor"></param>
		/// <param name="BackgroundColor"></param>
		/// <param name="AlternateBackgroundColor"></param>
		public void SetColors(Color ForegroundColor,Color BackgroundColor,Color AlternateBackgroundColor)
		{
			foreach (WordCell rc in this.Cells)	rc.SetForegroundColor(ForegroundColor);
			this.SetBackgroundColor(AlternateBackgroundColor);
			for (int r=this.startRow;r<=this.endRow;r+=2)
				for (int c=this.startColumn;c<=this.endColumn;c++)
					this.WordTable.Cell(r,c).SetBackgroundColor(BackgroundColor);
		}
		/// <summary>
		/// Sets those colors to each cell of the cellrange.
		/// </summary>
		/// <param name="ForegroundColor"></param>
		/// <param name="BackgroundColor"></param>
		public void SetColors(Color ForegroundColor,Color BackgroundColor)
		{
			foreach (WordCell rc in this.Cells)	rc.SetForegroundColor(ForegroundColor);
			foreach (WordCell rc in this.Cells)	rc.SetBackgroundColor(BackgroundColor);
		}
		/// <summary>
		/// Sets this content alignment to each cell of the cellrange.
		/// </summary>
		/// <param name="ContentAlignment"></param>
		public void SetContentAlignment(ContentAlignment ContentAlignment)
		{
			foreach (WordCell rc in this.Cells)	rc.contentAlignment=ContentAlignment;
		}
		/// <summary>
		/// Sets this font to each cell of the cellrange.
		/// </summary>
		/// <param name="Font"></param>
		public void SetFont(Font Font)
		{
			foreach (WordCell rc in this.Cells) rc.SetFont(Font);
		}
		/// <summary>
		/// Sets this border to the sub-table area defined by the cellrange
		/// </summary>
		/// <param name="BorderColor"></param>
		/// <param name="BorderWidth"></param>
		/// <param name="top"></param>
		/// <param name="bottom"></param>
		/// <param name="left"></param>
		/// <param name="right"></param>
		public void SetBorders(Color BorderColor,int BorderWidth,bool top,bool bottom,bool left,bool right)
		{
			if (top)
			{
				for (int c=this.startColumn;c<=this.endColumn;c++)
					this.WordTable.Cell(this.startRow,c).SetTopBorder(BorderColor,BorderWidth);
			}
			if (bottom)
			{
				for (int c=this.startColumn;c<=this.endColumn;c++)
					this.WordTable.Cell(this.endRow,c).SetBottomBorder(BorderColor,BorderWidth);
			}
			if (right)
			{
				for (int r=this.startRow;r<=this.endRow;r++)
					this.WordTable.Cell(r,this.endColumn).SetRightBorder(BorderColor,BorderWidth);
			}
			if (left)
			{
				for (int r=this.startRow;r<=this.endRow;r++)
					this.WordTable.Cell(r,this.startColumn).SetLeftBorder(BorderColor,BorderWidth);
			}
		}
		/// <summary>
		/// Merges together every cell of this cellrange. (using row and column spanning)
		/// </summary>
		public void MergeCells()
		{
			this.WordTable.Rows[startRow][startColumn].rowSpan=endRow-startRow+1;
			for (int r=this.startRow;r<this.endRow+1;r++)
				this.WordTable.Rows[r][startColumn].colSpan=endColumn-startColumn+1;
		}
	}
}
