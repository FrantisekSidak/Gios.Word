using System;
using System.IO;
using System.Drawing;
using System.Data;
using System.Collections;
using Gios.Word;

namespace RTF_Beta_Test
{
	class Class1
	{
		[STAThread]
		static void Main(string[] args)
		{
			Gios.Word.WordDocument rd=new WordDocument(WordDocumentFormat.Letter_8_5x11);
			
			Font bold=new Font("Helvetica",12,FontStyle.Bold);
			Font regular=new Font("Helvetica",12,FontStyle.Regular);
			rd.SetFont(bold);
			rd.SetTextAlign(WordTextAlign.Left);
			WordTable rt=rd.NewTable(regular,Color.Black,7,4,0);
			rt.SetColumnsWidth(new int[]{50,9,40,1});
			//foreach (WordCell rc in rt.Cells) rc.SetBorders(Color.Black,1,true,true,true,true);
			
			rt.SetContentAlignment(ContentAlignment.TopLeft);
			rt.Rows[0].SetRowHeight(300);
			rt.Rows[1].SetRowHeight(1400);
			rt.Rows[0][0].RowSpan=3;
			rt.Rows[0][0].SetContentAlignment(ContentAlignment.MiddleCenter);
			rt.Rows[0][0].PutImage(@"..\..\cp.jpg",70);
			rt.Rows[1][2].SetCellPadding(100);
			rt.Rows[1][2].SetContentAlignment(ContentAlignment.MiddleLeft);
			rt.Rows[1][2].SetFont(new Font("Helvetica",9,FontStyle.Bold));
			rt.Rows[1][2].WriteLine("GIOS PAOLO");
			rt.Rows[1][2].WriteLine("ELM STREET, 59");
			rt.Rows[1][2].Write("SPRINGFIELD");
			rt.Rows[1][2].SetBorders(Color.Black,1,true,true,true,true);
			;
			rt.Rows[4][0].SetFont(bold);
			rt.Rows[4][0].ColSpan=4;
			rt.Rows[4][0].WriteLine();
			rt.Rows[4][0].WriteLine("Gios Technologies - Power With Semplicity.\n\n");
			
			rt.Rows[5][0].WriteLine(DateTime.Today.ToLongDateString()+"\n\n\n\n");
			rt.Rows[5][1].ColSpan=3;
			rt.Rows[5][1].SetContentAlignment(ContentAlignment.TopRight);
			rt.Rows[5][1].WriteLine("Receipt Number 01302");
			
			WordCell body=rt.Rows[6][0];
			body.ColSpan=4;
			body.SetFont(bold);
			body.WriteLine("WORD .NET:");
			body.WriteLine();
			body.SetFont(regular);
			body.WriteLine("A complete library in c# for generating Word Documents using the Rich Text Format Specification!");
			body.WriteLine();
			body.WriteLine();
			body.SetFont(bold);
			body.Write("blah blah blah... ");
			body.SetFont(regular);
			body.WriteLine("blah blah blah\n");
			
			rt.SaveToDocument(10000,0);
			rd.SetPageNumbering(12);

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


			rd.SaveToFile("..\\..\\Example1.doc");
			System.Diagnostics.Process.Start("..\\..\\Example1.doc");
			
		}
		

	}
}
