/*
 * Created by SharpDevelop.
 * User: qingwa2009
 * Date: 2022/1/12
 * Time: 13:45
 * 
 * 关于设置哪些接口可见，哪些不可见的：
 * 类默认是都可见的，就是[ClassInterface(ClassInterfaceType.AutoDual)]
 * 类写上[ClassInterface(ClassInterfaceType.None)]就会不可见
 * 接口写上[InterfaceType(ComInterfaceType.InterfaceIsDual)]，
 * 然后类实现该接口就只有该接口的方法可见 
 */
using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Test
{
	[InterfaceType(ComInterfaceType.InterfaceIsDual)]
	[Guid("F6928DCD-B9C1-4D1C-A03B-EBB3E2BAFA18")]
	public interface IMyCom
	{
		[DispId(1)]
		string Text{get; set;}
		[DispId(2)]
		bool Visible{get; set;}
	}
	
	[ClassInterface(ClassInterfaceType.None)]	
	[Guid("1BFC9662-614E-40B8-9C55-4647464D346E")]
	public partial class Form1 : Form, IMyCom
	{
		public Form1()
		{
			//
			// The InitializeComponent() call is required for Windows Forms designer support.
			//
			InitializeComponent();
			
			//
			// TODO: Add constructor code after the InitializeComponent() call.
			//
			
		}
	}
}
