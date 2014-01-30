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
	internal class WordString : IWordStreamer
	{
		private string _value;
		public WordString(string Value,bool DoEncoding)
		{
			if (DoEncoding)
				this._value=Utility.Encode(Value);
			else
                this._value=Value;
		}

		#region IWordStreamer Members

		public void RenderToStream(System.IO.Stream ms)
		{
			Byte[] b=Utility.StringToByte(this._value);
			ms.Write(b,0,b.Length);
		}

		#endregion
	}
}
