using System;
using System.Drawing;
using System.Text;
using Gios.Word;

namespace Word_Example2
{
	class Class1
	{
		[STAThread]
		static void Main(string[] args)
		{
			// Create a new Word Document
			WordDocument rd=new WordDocument(WordDocumentFormat.Letter_8_5x11_Horizontal);

			// Sets a text effect using a custom control word. See RTF spefications for further implementations.
			rd.WriteControlWord("animtext3");

			// I think that there's no need of explanaition...
			rd.SetFontBackgroundColor(Color.Yellow);
			rd.SetForegroundColor(Color.Red);
			rd.SetFont(new Font("Arial",30));
			rd.SetTextAlign(WordTextAlign.Center);

			rd.WriteLine("GIOS WORD.NET Features Test");

			// Reset the text effect using a custom control word.
			rd.WriteControlWord("animtext0");
			
			// Resets the colors
			rd.SetFontBackgroundColor();
			rd.SetForegroundColor(Color.Black);

			rd.WriteLine();

			WordTable rt1=rd.NewTable(new Font("Arial",12),Color.Red,2,2,2);
			rt1.Rows[0][0].Write("hello");
			rt1.SetBorders(Color.Red,2,true,true,true,true);
			rt1.SaveToDocument(3000,4000);		
			rd.WriteLine();

			WordTable rt2=rd.NewTable(new Font("Arial",12),Color.Blue,2,4,2);
			rt2.Rows[0][0].Write("hello");
			rt2.SaveToDocument(3000,0);
			rd.WriteLine();

			// We can also set the starting page of the document.
			rd.SetPageNumbering(999);

			// In a object oriented document description, we can set the header
			// and the footer just we want to do it!
			rd.HeaderStart();
			rd.SetTextAlign(WordTextAlign.Center);
			rd.SetFont(new Font("Verdana",10,FontStyle.Bold));
			rd.Write("Paolo Gios, ICT Consultant - http://www.paologios.com");
			rd.WriteLine();
			rd.HeaderEnd();

			rd.FooterStart();
			rd.SetTextAlign(WordTextAlign.Center);
			rd.SetFont(new Font("Verdana",10,FontStyle.Bold));
			rd.Write("Copyright © 2005 by Paolo Gios - All rigths reserved");
			rd.FooterEnd();

			rd.NewPage();
			
			rd.SetTextAlign(WordTextAlign.Left);
			rd.WriteLine();
			rd.WriteLine("And this is a list using the line indent:");
			rd.WriteLine();
			// this sets the distance from the margin
			rd.SetParagraph(500);
			rd.WriteLine("- One");
			rd.WriteLine("- Two");
			rd.WriteLine("- Three");
			// resets the paragraph
			rd.SetParagraph();
			rd.WriteLine();
			rd.WriteLine("That's all... Enjoy!");
		
			rd.SaveToFile("..\\..\\Example2.doc");
			System.Diagnostics.Process.Start("..\\..\\Example2.doc");
			
			
		}
	}
}
