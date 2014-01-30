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
using System.IO;

namespace Gios.Word
{
	/// <summary>
	/// The target Word Document.
	/// </summary>
	public class WordDocument : WordArea
	{
		#region constructor and properties
		/// <summary>
		/// Creates a new word document specifing the format of the pages.
		/// </summary>
		/// <param name="WordDocumentFormat"></param>
		public WordDocument(WordDocumentFormat WordDocumentFormat)
		{
			this.WordDocument=this;
			this.FontList=new ArrayList();
			this.Colors=new ArrayList();
			this.AddColorAndGetID(Color.Black);
			this.AddColorAndGetID(Color.White);
			this.Objects=new ArrayList();
			this.Write(WordDocumentFormat.ToLineStream(),false);
		}
		private ArrayList FontList;
		private ArrayList Colors;
		private int StartPage=0;
		#endregion

		#region public methods
		/// <summary>
		/// Writes a page interruption (new page)
		/// </summary>
		public void NewPage()
		{
			this.Write("\n\\page\n",false);
		}
		internal WordImage NewImage(string file,int DPI)
		{
			return new WordImage(file,DPI);
		}
		/// <summary>
		/// Creates a new WordTable.
		/// </summary>
		/// <param name="DefaultFont"></param>
		/// <param name="DefaultForegroundColor"></param>
		/// <param name="rows"></param>
		/// <param name="columns"></param>
		/// <param name="padding"></param>
		/// <returns></returns>
		public WordTable NewTable(Font DefaultFont,Color DefaultForegroundColor,int rows,int columns,int padding)
		{
			WordTable ta=new WordTable(this.WordDocument,rows,columns,padding);
			foreach (WordCell rc in ta.cells.Values) 
			{
				rc.SetForegroundColor(DefaultForegroundColor);
				rc.SetFont(DefaultFont);
			}
			return ta;
		}
		/// <summary>
		/// Sets the document's page numbering. It can be called only once.
		/// </summary>
		/// <param name="StartPageNumber"></param>
		public void SetPageNumbering(int StartPageNumber)
		{
			if (StartPageNumber<1) throw new Exception("StartPageNumber must be greater than zero.");
			if (this.StartPage>0) throw new Exception("Page Numbering is already set.");
			this.StartPage=StartPageNumber;
		}
		/// <summary>
		/// Set the starting tag of the header of the document
		/// </summary>
		public void HeaderStart()
		{
			this.Write("{\\header ",false);
		}
		/// <summary>
		/// Set the ending tag of the header of the document
		/// </summary>
		public void HeaderEnd()
		{
			this.Write("\\par }",false);
		}
		/// <summary>
		/// Set the starting tag of the footer of the document
		/// </summary>
		public void FooterStart()
		{
			this.Write("{\\footer ",false);
		}
		/// <summary>
		/// Set the ending tag of the footer of the document
		/// </summary>
		public void FooterEnd()
		{
			this.Write("\\par }",false);
		}
		#endregion

		#region internal methods
		internal int AddFontAndGetID(Font f)
		{
			if (!FontList.Contains(f.Name))
			{
				FontList.Add(f.Name);
			}
			return FontList.IndexOf(f.Name);
		}
		internal int AddColorAndGetID(Color c)
		{
			string cl=Utility.ColorLine(c);
			if (!this.Colors.Contains(cl))
			{
				int id=this.Colors.Count;
				this.Colors.Add(cl);
				return id+1;
			} 
			return this.Colors.IndexOf(cl)+1;
		}
		#endregion

		#region calculated properties
		internal string Header
		{
			get
			{
				string s= "{";
				s+="\\rtf1\\ansi\n";
				s+=this.FontTable;
				s+=this.ColorTable;
				if (this.StartPage>-1) s+="\\pgnstart"+this.StartPage+"\n";
				return s;
			}
		}
		internal string Footer
		{
			get
			{
				return "\n}\n";
			}
		}
		
		internal string FontTable
		{
			get
			{
				string s="{\\fonttbl\n";
				for (int index=0;index<this.FontList.Count;index++)
				{
					s+="{\\f"+index.ToString()+" \\fcharset0 "+FontList[index].ToString()+";}\n";
				}
				s+="}\n";
				return s;
			}
		}
		
		internal string ColorTable
		{
			get
			{
				string s="{\\colortbl;\n";
				foreach (string c in this.Colors)
					s+=c;
				s+="}\n";
				return s;
			}
		}
		#endregion

		#region rendering
		
		/// <summary>
		/// Outputs the Complete WORD Document to a Generic Stream as a Rich Text Format (RTF)
		/// </summary>
		/// <param name="Stream">
		/// The Generic Stream to Output the Pdf Document
		/// </param>

		public void SaveToStream(System.IO.Stream Stream)
		{
			Utility.Send(this.Header,Stream);
			System.IO.MemoryStream ms2=new System.IO.MemoryStream();
			this.RenderToStream(ms2);
			Byte[] b=ms2.ToArray();
			Stream.Write(b,0,b.Length);
			ms2.Close();
			Utility.Send(this.Footer,Stream);
		}
		/// <summary>
		/// Outputs the complete WORD Document to a file as a Rich Text Format (RTF)
		/// </summary>
		/// <param name="file"></param>
		public void SaveToFile(string file)
		{
			System.IO.StreamWriter sw=null;
			try
			{
				sw=new StreamWriter(file,false);
				this.SaveToStream(sw.BaseStream);
			
			}
			catch
			{
				throw new Exception("Error opening destination file");
			}
			finally
			{
				sw.Close();
			}
		}
		#endregion
	}
}
