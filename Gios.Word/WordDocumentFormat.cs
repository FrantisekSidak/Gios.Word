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

namespace Gios.Word
{
	/// <summary>
	/// Summary description for WordDocumentFormat.
	/// </summary>
	public class WordDocumentFormat
	{
		private int width,height;
		private int margl,margr,margt,margb;
		
		private WordDocumentFormat()
		{
			
		}
		/// <summary>
		/// gets the Classic A4 Letter.
		/// </summary>
		public static WordDocumentFormat A4
		{
			get
			{
				return WordDocumentFormat.InCentimeters(21,29.7,2,2,2.5,2);
			}
		}
		/// <summary>
		/// gets the Classic A4 Letter (Horizontal)
		/// </summary>
		public static WordDocumentFormat A4_Horizontal
		{
			get
			{
				return WordDocumentFormat.InCentimeters(29.7,21,2,2.5,2,2);
			}
		}
		/// <summary>
		/// gets the 8.5x11 American Letter.
		/// </summary>
		public static WordDocumentFormat Letter_8_5x11
		{
			get
			{
				return WordDocumentFormat.InInches(8.5,11,.75,.75,1,.75);
			}
		}
		/// <summary>
		/// gets the 8.5x11 American Letter. (Horizontal)
		/// </summary>
		public static WordDocumentFormat Letter_8_5x11_Horizontal
		{
			get
			{
				return WordDocumentFormat.InInches(11,8.5,.75,1,.75,.75);
			}
		}
		/// <summary>
		/// creates a centimeters custom sized paper.
		/// </summary>
		/// <param name="Width"></param>
		/// <param name="Height"></param>
		/// <param name="LeftMargin"></param>
		/// <param name="RightMargin"></param>
		/// <param name="TopMargin"></param>
		/// <param name="BottomMargin"></param>
		/// <returns></returns>
		public static WordDocumentFormat InCentimeters(double Width,double Height,double LeftMargin,double RightMargin,double TopMargin,double BottomMargin)
		{
			WordDocumentFormat rpf=new WordDocumentFormat();
			rpf.height=(int)(Height*16840/29.7);
			rpf.width=(int)(Width*16840/29.7);
			rpf.margl=(int)(LeftMargin*16840/29.7);
			rpf.margr=(int)(RightMargin*16840/29.7);
			rpf.margt=(int)(TopMargin*16840/29.7);
			rpf.margb=(int)(BottomMargin*16840/29.7);
			
			return rpf;
		}
		/// <summary>
		/// creates a inches custom sized paper.
		/// </summary>
		/// <param name="Width"></param>
		/// <param name="Height"></param>
		/// <param name="LeftMargin"></param>
		/// <param name="RightMargin"></param>
		/// <param name="TopMargin"></param>
		/// <param name="BottomMargin"></param>
		/// <returns></returns>

		public static WordDocumentFormat InInches(double Width,double Height,double LeftMargin,double RightMargin,double TopMargin,double BottomMargin)
		{
			return WordDocumentFormat.InCentimeters(
				Width*2.54,Height*2.54,LeftMargin*2.54
				,RightMargin*2.54,TopMargin*2.54,BottomMargin*2.54);
		}
		internal string ToLineStream()
		{
			string s="\\paperw"+this.width+"\\paperh"+this.height
				+"\\margl"+this.margl+"\\margr"+this.margr
				+"\\margt"+this.margt+"\\margb"+this.margb;
			if (width>height) s+="\\landscape1";
			s=s+"\n";
			return s;
		}
	}
}
