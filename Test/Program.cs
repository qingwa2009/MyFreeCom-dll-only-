/*
 * Created by SharpDevelop.
 * User: qingwa2009
 * Date: 2021/12/15
 * Time: 13:42
 * 
 * 免管理员权限注册互操作COM
 * 
 * [assembly: ComVisible(true)] 互操作这一句是必须的！
 * 
 * VBA测试时注意COM被VBA占用而导致COM编译失败的问题，关掉Excel就可以了
 * 
 */
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Test
{
	
	
	class Program
	{
		[STAThread]
		public static void Main(string[] args)
		{
			
			int dwRegister;

			Form1 f= new Form1();
			if(MyFreeCom.MainFunc.AddToROT(f, out dwRegister)){
				MyFreeCom.MainFunc.RegisterMyCom(f, "MyClass.Application");							
				Application.Run(f);							
				MyFreeCom.MainFunc.UnregisterMyCom(f, "MyClass.Application");				
				MyFreeCom.MainFunc.RemoveFromROT(dwRegister);
			}
			
			
		}
	}
}