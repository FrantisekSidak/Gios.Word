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
using System.Text;
using System.Collections;
using System.Drawing;

namespace Gios.Word
{
	/// <summary>
	/// cell of a WordTable
	/// </summary>
	public class WordCell : WordArea
	{
		#region fields
		private bool tb,bb,lb,rb;
		private int bbWidth,lbWidth,rbWidth,tbWidth,lp=0,rp=0,bp=0,tp=0;
		private Color bgColor,tbColor,bbColor,lbColor,rbColor;
		internal ContentAlignment contentAlignment;
		internal int columnIndex,rowIndex;
		internal bool hasBackgroundColor;
		internal int colSpan=1,rowSpan=1;
		internal WordTable WordTable;
		#endregion

		#region readonly properties
		internal bool HasPadding
		{
			get
			{
				return lp>0 || bp>0 || tp>0 || rp>0;
			}
		}
		internal bool isRowSpanned
		{
			get
			{
				for (int row=0;row<this.rowIndex;row++)
					if (this.WordTable.Rows[row][this.columnIndex].rowSpan+row>this.rowIndex)
						return true;
				return false;
			}
		}
		internal bool isRowSpanner
		{
			get
			{
				if (this.rowSpan>1) return true;
				return false;
			}
		}
		internal bool isColumnSpanned
		{
			get
			{
				for (int col=0;col<this.columnIndex;col++)
					if (this.WordTable.Rows[this.rowIndex][col].colSpan+col>this.columnIndex)
						return true;
				return false;
			}
		}
		internal bool isColumnSpanner
		{
			get
			{
				if (this.colSpan>1) return true;
				return false;
			}
		}
		internal string PaddingString
		{
			get
			{
				StringBuilder sb=new StringBuilder();
				sb.Append("\\clpadl");
				sb.Append(this.lp);
				sb.Append("\\clpadt");
				sb.Append(this.tp);
				sb.Append("\\clpadr");
				sb.Append(this.rp);
				sb.Append("\\clpadb");
				sb.Append(this.bp);
				sb.Append("\\clpadfl3\\clpadft3\\clpadfb3\\clpadfr3");
				return sb.ToString();
			}
		}
		internal string BorderString
		{
			get
			{
				string s="";
				if ((this.bb)||(this.tb)||(this.rb)||(this.lb))
				{
					string m="\\clbrdr*\\brdrw+\\brdrs\\brdrcf";
					if (this.rb) 
						s+=m.Replace("*","r").Replace("+",this.rbWidth.ToString())
							+this.WordDocument.AddColorAndGetID(this.rbColor);
					if (this.tb) 
						s+=m.Replace("*","t").Replace("+",this.tbWidth.ToString())
							+this.WordDocument.AddColorAndGetID(this.tbColor);
					if (this.bb)
						s+=m.Replace("*","b").Replace("+",this.bbWidth.ToString())
							+this.WordDocument.AddColorAndGetID(this.bbColor);
					if (this.lb)
						s+=m.Replace("*","l").Replace("+",this.lbWidth.ToString())
							+this.WordDocument.AddColorAndGetID(this.lbColor);
				}
				return s;
			}
		}
		#endregion
	
		#region internal properties
		internal Color BackgroundColor
		{
			get
			{
				return this.bgColor;
			}
			set
			{
				this.hasBackgroundColor=true;
				this.bgColor=value;
			}
		}
		#endregion
		
		#region public properties
		/// <summary>
		/// Gets or sets the column span of the cell.
		/// </summary>
		public int ColSpan
		{
			get
			{
				return this.colSpan;
			}
			set
			{
				this.colSpan=value;
			}
		}
		/// <summary>
		/// Gets or sets the row span of the cell.
		/// </summary>
		public int RowSpan
		{
			get
			{
				return this.rowSpan;
			}
			set
			{
				this.rowSpan=value;
			}
		}
		#endregion
		
		#region public methods
		/// <summary>
		/// Sets the padding of the cell.
		/// </summary>
		/// <param name="Padding"></param>
		public void SetCellPadding(int Padding)
		{
			lp=Padding;
			rp=Padding;
			tp=Padding;
			bp=Padding;
		}
		/// <summary>
		/// Sets the background color of the cell.
		/// </summary>
		/// <param name="BackgroundColor"></param>
		public void SetBackgroundColor(Color BackgroundColor)
		{
			this.BackgroundColor=BackgroundColor;
		}
		/// <summary>
		/// Sets the content alignment of this cell.
		/// </summary>
		/// <param name="ContentAlignment"></param>
		public	void SetContentAlignment(ContentAlignment ContentAlignment)
		{
			this.contentAlignment=ContentAlignment;
		}
		internal WordCell(WordDocument WordDocument,WordTable WordTable,int rowIndex,int columnIndex)
		{
			this.Objects=new ArrayList();
			this.contentAlignment=ContentAlignment.MiddleCenter;
			this.WordDocument=WordDocument;
			this.WordTable=WordTable;
			this.columnIndex=columnIndex;
			this.rowIndex=rowIndex;
		}
		internal WordCell Traslate(int x,int y)
		{
			return this.WordTable.Cell(this.rowIndex+y,this.columnIndex+x);
		}
		/// <summary>
		/// Sets each border of the cell.
		/// </summary>
		/// <param name="BorderColor"></param>
		/// <param name="BorderWidth"></param>
		/// <param name="top"></param>
		/// <param name="bottom"></param>
		/// <param name="left"></param>
		/// <param name="right"></param>
		public void SetBorders(Color BorderColor,int BorderWidth,bool top,bool bottom,bool left,bool right)
		{
			if ( this.isColumnSpanner || this.isRowSpanner )
			{
				this.WordTable.CellRange(this.rowIndex,this.columnIndex,this.rowIndex+this.rowSpan-1,
					this.columnIndex+this.colSpan-1).SetBorders(BorderColor,BorderWidth,top,bottom,left,right);
			}
			else
			{
				if (top) this.SetTopBorder(BorderColor,BorderWidth);
				if (bottom) this.SetBottomBorder(BorderColor,BorderWidth);
				if (left) this.SetLeftBorder(BorderColor,BorderWidth);
				if (right) this.SetRightBorder(BorderColor,BorderWidth);
			}
		}
		internal void SetTopBorder(Color TopBorderColor,int TopBorderWidth)
		{
			
			this.tbColor=TopBorderColor;
			this.tbWidth=TopBorderWidth;
			this.tb=true;
		}
		internal void SetBottomBorder(Color BottomBorderColor,int BottomBorderWidth)
		{
			WordCell rc=this.Traslate(0,this.rowSpan);
			if (rc!=null)
			{
				rc.tbColor=BottomBorderColor;
				rc.tbWidth=BottomBorderWidth;
				rc.tb=true;
			}
			if (!this.isRowSpanned)
			{
				this.bbColor=BottomBorderColor;
				this.bbWidth=BottomBorderWidth;
				this.bb=true;
			}
			else
			{
				this.Traslate(0,-1).SetBottomBorder(BottomBorderColor,BottomBorderWidth);
			}
		}
		internal void SetLeftBorder(Color LeftBorderColor,int LeftBorderWidth)
		{
			this.lbColor=LeftBorderColor;
			this.lbWidth=LeftBorderWidth;
			this.lb=true;
		}
		internal void SetRightBorder(Color RightBorderColor,int RightBorderWidth)
		{
			WordCell rc=this.Traslate(this.colSpan,0);
			if (rc!=null)
			{
				rc.lbColor=RightBorderColor;
				rc.lbWidth=RightBorderWidth;
				rc.lb=true;
			}
			if (!this.isColumnSpanned)
			{
				this.rbColor=RightBorderColor;
				this.rbWidth=RightBorderWidth;
				this.rb=true;
			} else
			 this.Traslate(-1,0).SetRightBorder(RightBorderColor,RightBorderWidth);
		}
		
		
		#endregion
		
	}
}
