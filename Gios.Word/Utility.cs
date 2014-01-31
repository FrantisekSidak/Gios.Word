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
using System.Text;

namespace Gios.Word
{
    internal class Utility
    {
        internal static string ColorLine(Color Color)
        {
            return "\\red" + (int)Color.R + "\\green" + (int)Color.G + "\\blue" + (int)Color.B + ";\n";
        }

        internal static string EncodeCellAlignV(ContentAlignment cellTextAlign)
        {
            switch (cellTextAlign)
            {
                case ContentAlignment.MiddleLeft:
                case ContentAlignment.MiddleCenter:
                case ContentAlignment.MiddleRight:
                    return "\\clvertalc";

                case ContentAlignment.BottomCenter:
                case ContentAlignment.BottomLeft:
                case ContentAlignment.BottomRight:
                    return "\\clvertalb";
            }
            return "\\clvertalt";
        }

        internal static string EncodeCellAlignH(ContentAlignment cellTextAlign)
        {
            switch (cellTextAlign)
            {
                case ContentAlignment.BottomCenter:
                case ContentAlignment.MiddleCenter:
                case ContentAlignment.TopCenter:
                    return "\\qc";

                case ContentAlignment.BottomRight:
                case ContentAlignment.MiddleRight:
                case ContentAlignment.TopRight:
                    return "\\qr";
            }
            return "\\ql";
        }

        static char[] hexDigits = {
									  '0', '1', '2', '3', '4', '5', '6', '7',
									  '8', '9', 'A', 'B', 'C', 'D', 'E', 'F'};

        internal static string ToHexString(byte[] bytes)
        {
            char[] chars = new char[bytes.Length * 2];
            for (int i = 0; i < bytes.Length; i++)
            {
                int b = bytes[i];
                chars[i * 2] = hexDigits[b >> 4];
                chars[i * 2 + 1] = hexDigits[b & 0xF];
            }
            return new string(chars);
        }

        internal static Byte[] StringToByte(string s)
        {
            return System.Text.ASCIIEncoding.ASCII.GetBytes(s);
        }

        internal static void Send(string strMsg, System.IO.Stream ms)
        {
            Byte[] buffer = null;
            buffer = System.Text.ASCIIEncoding.ASCII.GetBytes(strMsg);
            ms.Write(buffer, 0, buffer.Length);
        }

        internal static string Encode(string s)
        {
            s = s.Replace("\\", "\\\\");
            var sb = new StringBuilder();

            foreach (var c in s)
            {
                if (c == '\\' || c == '{' || c == '}')
                    sb.Append(@"\" + c);
                else if ((c >= 48 && c <= 122) || c == '\n' || c == ' ')
                    sb.Append(c);
                else if (c <= 0x7f)
                    sb.Append(c);
                else
                    sb.Append("\\u" + Convert.ToUInt32(c) + "?");
            }

            return sb.ToString().Replace("\r\n", "\\par").Replace("\n", "\\par ");
        }
    }
}
