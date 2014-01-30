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
	/// the different kinds of text alignment.
	/// </summary>
	public enum WordTextAlign
	{
		/// <summary>
		/// left aligned text.
		/// </summary>
		Left=0,
		/// <summary>
		/// center aligned text.
		/// </summary>
		Center=1,
		/// <summary>
		/// right aligned text.
		/// </summary>
		Right=2,
		/// <summary>
		/// jusified text.
		/// </summary>
		Justified=3
	}
}
