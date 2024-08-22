/*
 * Created by SharpDevelop.
 * User: Rice Cipriani
 * Date: 07/05/2012
 * Time: 21:54
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;

namespace GobboManuale
{
	/// <summary>
	/// Description of Class1.
	/// </summary>
	public class Line
	{
        private int nLine;
		private int start;
		private string text;
		public Line(int start, string text, int nLine)
		{
            this.nLine = nLine;
			this.start=start;
			this.text=text;
		}
        public int getNLine()
        {
            return nLine;
        }
		public int getStart(){
			return start;
		}
		public string getText(){
			return text;
		}
	}
}
