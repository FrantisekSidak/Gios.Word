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
using System.Text;

namespace Gios.Word
{
	/// <summary>
	/// the abstract class of an area of the document. Could be a Table Cell or the Document itself.
	/// </summary>
	public abstract class WordArea : IWordStreamer
	{
		#region public properties and constructors
		internal ArrayList Objects;
		internal WordTextAlign currentAlign;
		internal Color currentColor;
		internal Font currentFont;		
		internal WordDocument WordDocument;
		#endregion

		#region constructors
		internal WordArea(WordDocument WordDocument)
		{
			this.Objects=new ArrayList();
			this.WordDocument=WordDocument;
		}
		internal WordArea()
		{
			this.Objects=new ArrayList();
		}
		#endregion

		#region internal methods
		
		internal void Write(string text,bool DoEncoding)
		{
			this.Objects.Add(new WordString(text,DoEncoding));					
		}

		
		#endregion

		#region public methods
		/// <summary>
		/// Sets the Font for the next writing.
		/// </summary>
		/// <param name="Font"></param>
		public void SetFont(Font Font)
		{
			this.WriteControlWord("f"+this.WordDocument.AddFontAndGetID(Font).ToString());
			this.WriteControlWord("fs"+((int)(Font.Size*2)).ToString());
			if (Font.Bold) this.WriteControlWord("b"); else this.WriteControlWord("b0");
			if (Font.Italic) this.WriteControlWord("i"); else this.WriteControlWord("i0");
			if (Font.Underline) this.WriteControlWord("ul"); else this.WriteControlWord("ulnone");
			this.currentFont=Font;
			//this.Write(" ",false);
		}
		/// <summary>
		/// Sets the color for the next writing.
		/// </summary>
		/// <param name="Color"></param>
		public void SetForegroundColor(Color Color)
		{
			this.WriteControlWord("cf"+this.WordDocument.AddColorAndGetID(Color));
			this.currentColor=Color;
		}
		/// <summary>
		/// Sets the color for the next writing.
		/// </summary>
		/// <param name="Color"></param>
		public void SetFontBackgroundColor(Color Color)
		{
			this.WriteControlWord("chcbpat"+this.WordDocument.AddColorAndGetID(Color));
		}
		/// <summary>
		/// Resets the font background color
		/// </summary>
		public void SetFontBackgroundColor()
		{
			this.WriteControlWord("chcbpat0");
		}
		
		/// <summary>
		/// Expert users only. Insert a control word into the Document. (see RTF specifcations)
		/// 
		/// For example:
		/// myDocument.WriteControlWord("animtext3");
		/// Starts sparkling text writing.
		/// </summary>
		/// <param name="ControlWord"></param>
		public void WriteControlWord(string ControlWord)
		{
			this.Write("\\"+ControlWord+" ",false);
		}
		/// <summary>
		/// Expert users only. Writes RTF code directly into the Document.
		/// This can be useful for implementing further RTF capabilities.
		/// </summary>
		/// <param name="RtfText"></param>
		public void WriteRtf(string RtfText)
		{
			this.Write(RtfText,false);
		}
		/// <summary>
		/// Writes a text line.
		/// </summary>
		/// <param name="text"></param>
		public void WriteLine(string text)
		{
			this.Write(text+"\n");
		}
		/// <summary>
		/// Writes a TAB (Tabulation) Space
		/// </summary>
		public void WriteTabSpace()
		{
			this.WriteControlWord("tab");
		}
		/// <summary>
		/// Sets the line Indent
		/// </summary>
		/// <param name="Indent"></param>
		public void SetParagraph(int Indent)
		{
			this.WriteControlWord("li"+Indent);
		}
		/// <summary>
		/// Sets the line Indent
		/// </summary>
		/// <param name="Indent"></param>
		/// <param name="FirstIndent"></param>
		public void SetParagraph(int Indent,int FirstIndent)
		{
			this.WriteControlWord("fi"+FirstIndent);
			this.WriteControlWord("li"+Indent);
		}
		/// <summary>
		/// Resets the line indent
		/// </summary>
		public void SetParagraph()
		{
			this.WriteControlWord("lin0\n");
		}
		/// <summary>
		/// Writes a new line.
		/// </summary>
		public void WriteLine()
		{
			this.Write("\n");
		}
		/// <summary>
		/// Writes a string
		/// </summary>
		/// <param name="text"></param>
		public void Write(string text)
		{
			this.Write(text,true);
		}
		/// <summary>
		/// Puts a jpeg image into the current position.
		/// </summary>
		/// <param name="JpegFile">the color RGB jpeg file.</param>
		/// <param name="DPI">target picture resolution</param>
		public void PutImage(string JpegFile,int DPI)
		{
			this.Objects.Add(this.WordDocument.NewImage(JpegFile,DPI));
		}
		/// <summary>
		/// Sets the horizontal alignment of the text.
		/// </summary>
		/// <param name="WordTextAlignment"></param>
		public void SetTextAlign(WordTextAlign WordTextAlignment)
		{
			this.currentAlign=WordTextAlignment;
			switch (WordTextAlignment)
			{
				case WordTextAlign.Center:
					this.WriteControlWord("qc");
					break;
				case WordTextAlign.Right:
					this.WriteControlWord("qr");
					break;
				case WordTextAlign.Justified:
					this.WriteControlWord("qj");
					break;
				default:
					this.WriteControlWord("ql");
					break;
			}
		}
		#endregion
		
		#region IWordStreamer Members
		/// <summary>
		/// 
		/// </summary>
		/// <param name="ms"></param>
		public void RenderToStream(System.IO.Stream ms)
		{
			foreach (IWordStreamer iws in this.Objects)
			{
				iws.RenderToStream(ms);
			}	
		}

		#endregion
	}
}
