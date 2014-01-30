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
using System.IO;
using System.Drawing.Imaging;


namespace Gios.Word
{
	internal class WordImage : IWordStreamer
	{
		private int dpi;
		System.IO.Stream stream;
		string file;

		internal int DPI
		{
			get
			{
				return this.dpi;
			}
		}
		internal WordImage(string file,int DPI)
		{
			this.dpi=DPI;
			try
			{
				Image image=Image.FromFile(file);
				image.Dispose();
			}
			catch
			{
				throw new Exception("Error opening the jpeg file");
			}
			this.file=file;
		}
		

		#region IWordStreamer Members

		public void RenderToStream(Stream ms)
		{
			this.stream=ms;
			Utility.Send("{\\pict\\jpegblip",ms);
			Utility.Send("\\picscalex"+7200/this.dpi,ms);
			Utility.Send("\\picscaley"+7200/this.dpi+"\n",ms);
				
			FileStream fs = File.OpenRead(this.file);
			byte[] data = new byte[fs.Length];
			fs.Read (data, 0, data.Length);
				
			Utility.Send(Utility.ToHexString(data).Replace("FF","FF\n"),ms);
			fs.Close();
				
			Utility.Send("\n}\n",ms);
		}

		#endregion
	}
}

