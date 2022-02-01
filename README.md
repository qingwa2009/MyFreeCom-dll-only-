# MyFreeCom-dll-only-
**免管理员权限注册COM组件**

该项目可以让任意C#的类对象可以被VBA的GetObjectj及CreateObject调用；

并且不需要管理员权限（必须能修改Current User注册表项，一般情况都是能改的，不能改那就没辙了；有管理员权限的就忽略吧！），win7 win10都有效，其他自测！

在MyFreeCom(dll only)文件夹里仅提供了MyFreeCom.dll文件;

使用方法：

将MyFreeCom.dll添加到项目的引用中；

调用RegisterMyCom注册需要互操作的对象，调用后可以被CreateObject创建对象；

调用AddToROT将对象添加进运行表，调用后可以被GetObject获取对象；

具体用法可以参考Test文件夹的Program.cs文件，vba可以参考Test文件夹下的vbatest文件。


MyFreeCom.dll仅提供4个函数：


**public static bool RegisterMyCom(object instance, string progName);**

````
/// <summary>
/// 将对象添加进CurrentUser注册表，
/// </summary>
/// <param name="obj">对象</param>
/// <param name="progName">对象名称如“Excel.Application”，在vba中GetObject(,progName)可以获取到对象</param>
/// <returns></returns>
````
**public static bool UnregisterMyCom(object obj, string progName);**
````
/// <summary>
/// 从CurrentUser注册表中移除，
/// 一般不用调用，只是把写进注册表的内容删掉而已，每次打开COM程序都会重新写一遍的，
/// 如果清了注册表，那CreateObject是无法创建对象的。
/// </summary>
/// <param name="obj">对象</param>
/// <param name="progName">对象名称</param>
/// <returns></returns>
````
**public static bool AddToROT(object instance, out int dwRegister);**
````
/// <summary>
/// 把对象添加进运行对象表;
/// 添加进运行对象表后才能被GetObject(,"My.Application")获取到
/// "My.Application"这个值就是RegisterMyCom第二个参数自己提供的值
/// </summary>
/// <param name="c">对象</param>
/// <param name="dwRegister">该变量在RemoveFromROT需要用到</param>
/// <returns></returns>
````
**public static void RemoveFromROT(int dwRegister);**
````
/// <summary>
/// 从运行对象表中移除
/// COM运行结束时调用，从运行对象表中移除
/// </summary>
/// <param name="dwRegister">从AddToROT获取的值</param>
````

该项目仅提供dll文件，一般情况够用了；

想要源码学习下的可以加下微信(qingwa2009appx)，备注写上“MyFreeCom源码”，作者厚颜无耻的讨点红包😝，包多少看心意就行；

源码也没啥东西，也是那四个函数，还有一下互操作的注意事项：

如：GUID要不要写，如何暴露指定接口，如何操作UI线程等等；

作者花费好几天收集总结的，太深入的东西作者就不太懂了😝，作者只是业余的，只是有vba调用C#对象的需求，花了些时间搞了下，基本够用就行。

项目用SharpDevelop编写，.Net Frameworks2.0

