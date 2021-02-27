**中文版本介绍**

# PyComCAD介绍及开发方法

## 1.综述
​		提到Autocad在工业界的二次开发，VB或者Lisp可能作为常用的传统的编程语言。但是，Python语言简洁，优雅，学习门槛低，理应在Autocad二次开发中占有一席之地，加上Python丰富而强大的第三方库，更让Python对于Autocad二次开发任务如虎添翼，使得快速开发出符合工程师自身需求的功能成为可能。Pycomcad恰恰就是的获取Autocad API的接口库。

​		Pycomcad的底层库是`win32com`和`pythoncom`，其中，win32com负责获取Autocad的接口，包括一些枚举值，pythoncom主要负责进行数据类型转换。Pycomcad设计理念非常的简单，就是把win32com中多层调用的函数或者属性包裹为一个函数，用以方便记忆调用以及减少敲码次数，而不用每次都按照AutoCAD对象模型树一层一层的调用。

​		当涉及到Autocad中特定对象的方法或者属性，建议查看本仓库下的`acadauto.chm`文件。

##  2.底层库安装

`pip install pywin32`将会安装win32com和pythoncom库。

##  3.简单小例子

```python
#准备工作
import sys
import win32com.client
import math
sys.path.append(r'D:\programming\pycomcad\PycomCAD') #这里填写pycomcad.py所在的文件夹
from pycomcad import *
acad=Autocad()  #打开Autocad，如果已有打开的Autocad，则对该Autocad进行连接
if not acad.IsEarlyBind: #判断是否是EarlyBind，如果不是则打开Earlybind模式
    acad.TurnOnEarlyBind() 
acad.ShowLineweight(True)   #设置打开线宽
#进行绘制
line=acad.AddLine(Apoint(0,0),Apoint(100,0)) #绘制线
circle=acad.AddCircle(Apoint(100,0),10)  #绘制圆
circleBig=acad.AddCircle(Apoint(0,0),110)
circleInner=acad.AddCircle(Apoint(0,0),90)
for i in range(15):
    angle=math.radians(24)
    line.Copy()
    circle.Copy()
    line.Rotate(Apoint(0,0),angle)  #对线和圆进行复制并旋转
    circle.Rotate(Apoint(0,0),angle)
text=acad.AddText('Code makes a better world!',Apoint(0,0),20)  #绘制文字
text.Alignment=win32com.client.constants.acAlignmentTopCenter  #设置文字的alignment方向
text.Move(Apoint(0,0),Apoint(0,-150)) #移动文字
underLine=acad.AddLine(Apoint(-175,-180),Apoint(175,-180)) #绘制下划线
underLine.Lineweight=win32com.client.constants.acLnWt030 
underLine.Copy()
underLine.Offset(5) #向上偏移下划线
# 保存文件
acad.SaveAsFile(r'pycomcad.dwg')
```

绘制出如下图形：

<img src="https://cdn.jsdelivr.net/gh/JohnYang1210/bloggitpic/img/20210226220039.png" alt="image-20210226220032506" style="zoom:50%;" />

## 4.Pycomcad的基本架构

### 4.1 指定特定版本

​		如果个人电脑上装有多个版本的Autocad，我们想针对特定版本的Autocad进行二次开发，只需要修改`pycomcad.py`中`win32com.client.Dsipatch`函数中的`ProgID`就可以了(比如要指定Autocad2016版本，只需要修改为`self.acad=win32com.client.Dispatch('AutoCAD.Application.20')`，本仓库默认为`Autocad.Application`:

```python
class Autocad:
	def __init__(self):
		try:
			self.acad=win32com.client.Dispatch(`ProgID`)  #修改此处的ProgID
			self.acad.Visible=True 
		except:
			Autocad.__init__(self)
```

Autocad版本号与ProgID对应关系表如下：

| AutoCAD Production | ProgID                   |
| ------------- | ------------------------ |
| AutoCAD 2004  | AutoCAD.Application.16   |
| AutoCAD 2005  | AutoCAD.Application.16.1 |
| AutoCAD 2006  | AutoCAD.Application.16.2 |
| AutoCAD 2007  | AutoCAD.Application.17   |
| AutoCAD 2008  | AutoCAD.Application.17.1 |
| AutoCAD 2009  | AutoCAD.Application.17.2 |
| AutoCAD 2010  | AutoCAD.Application.18   |
| AutoCAD 2011  | AutoCAD.Application.18.1 |
| AutoCAD2014   | AutoCAD.Application.19   |
| AutoCAD2016   | AutoCAD.Application.20   |



### 4.2 模块级函数

* Apoint :点函数，传入x,y,z（可选，默认为0），其返回值作为其他函数的输入值。如在上面的例子中，`AddLine`与`AddCircle`方法均需输入Apoint函数的返回值
* ArrayTransform:将任何型式的数组转换为所需要的实数型数组，多用于pycomcad模块内部使用
* VtVertex：将分散的数据转换为实数型数组，多用于pycomcad模块内部使用
* VtFloat:将仅有数字的列表转换为实数型列表
* VtInt:将仅有整数的列表转换为整数型列表
* VtVariant:将变量型列表转换为变量型列表
* AngleDtoR:将°转换为radian，也可以用math.radians
* AngleRtoD:将radian转换为°，也可以用math.degrees
* FilterType，FilterData，过滤规则中将DXF码传入，也可以用VtInt(ft),VtVariant(fd)，详细参见`GetSelectionSets`方法说明

有关数据转换更详细的信息见： https://www.cnblogs.com/johnyang/p/12617881.html .

## 4.3 Early-bind模式还是Lazy-bind模式

​		在上面例子中，我们用到了`acad.IsEarlybind`属性，以及`acad.TurnOnEarlyBind`方法，那么什么是early-bind模式，什么是lazy-bind模式呢？

​		默认地，pycomcad是lazy-bind模式，意思就是pycomcad对于特定的对象，比如线，圆等的方法，属性，以及常量枚举值，事先是不知道的，而early-bind模式下，pycomcad就提前知道了特定的对象的类型，方法，属性。实际上，这对于我们进行二次开发是有比较大的影响的，因为有时候我们需要知道选中的对象到底是什么样的类型，然后根据其类型，进行不同的操作。比如，对于early-bind模式，pycomcad能识别`win32com.client.constants.acRed`枚举值，而lazy-bind模式下，不能对其进行识别。建议把early-bind模式打开。

​		Autocad对象，比如它是`acad`，它的`IsEarlyBind`属性可以判断目前Autocad的模式是哪一种，如果是earlyy-bind模式，则返回`True`,否则返回`False`，如果是lazy-bind，那么可以调用`TrunOnEarlyBind()`方法来转变为Early-bind模式。

​		有关Early-bin和Lazy-bind模式的信息，详见我的博文：https://www.cnblogs.com/johnyang/p/12521301.html

## 4.4 模块的主要方法及属性

------



* 系统变量

| 方法        | 作用         |
| ----------- | ------------ |
| SetVariable | 设置系统变量 |
| GetVariable | 获取系统变量 |

------



* 文件处理

| 方法          | 作用                             |
| ------------- | -------------------------------- |
| OpenFile      | 打开文件                         |
| CreateNewFile | 新建文件                         |
| SaveFile      | 保存                             |
| SaveAsFile    | 将文件保存至设定路径下           |
| Close         | 关闭文件                         |
| PurgeAll      | Purge文件                        |
| Regen         | Regen文件                        |
| GetOpenedFile | 返回指定的已经打开的文件         |
| ActivateFile  | 指定已经打开的文件为当前工作文件 |
| DeepClone     | 跨文件间复制对象                 |

| 属性              | 作用                                      |
| ----------------- | ----------------------------------------- |
| OpenedFilenames   | 返回已经打开的所有文件列表                |
| OpenedFilenumbers | 返回已经打开的文件的数量                  |
| CurrentFilename   | 返回当前文件名                            |
| FilePath          | 返回当前文件路径                          |
| IsSaved           | 如果当前文件保存了，则返回True，否则False |

------



* 精细绘图设置

| 方法        | 作用              |
| ----------- | ----------------- |
| ZoomExtents | 极限放大          |
| ZoomAll     | 显示整个图形      |
| GridOn      | 打开/关闭栅格     |
| SnapOn      | 打开/关闭捕捉状态 |

------



* 创建实体

| 方法       | 作用                  |
| ---------- | --------------------- |
| AddPoint   | 创建点                |
| AddLine    | 创建线                |
| AddLwpline | 创建LightWeight多段线 |
| AddCircle  | 创建圆                |
| AddArc     | 创建圆弧              |
| AddTable   | 创建表                |
| AddSpline  | 创建拟合曲线          |
| AddEllipse | 创建椭圆              |
| AddHatch   | 创建填充              |
| AddSolid   | 创建实心面            |

------



* 引用及选择

| 方法             | 作用                                                         |
| ---------------- | ------------------------------------------------------------ |
| Handle2Object    | 通过实体引用的Handle值获取实体本身                           |
| GetEntityByItem  | 通过实体的索引来获取实体                                     |
| GetSelectionSets | 获取选择集（选择及过滤机制详见https://www.cnblogs.com/johnyang/p/12934674.html） |

------



* 图层

| 方法          | 作用         |
| ------------- | ------------ |
| CreateLayer   | 创建图层     |
| ActivateLayer | 激活图层     |
| GetLayer      | 获取图层对象 |

| 属性         | 作用           |
| ------------ | -------------- |
| LayerNumbers | 返回图层的总数 |
| LayerNames   | 返回图层名列表 |
| Layers       | 返回图层集     |
| ActiveLayer  | 返回当前图层   |

------



* 线型

| 方法             | 作用          |
| ---------------- | ------------- |
| LoadLinetype     | 加载线型      |
| ActivateLinetype | 激活线型      |
| ShowLineweight   | 显示/关闭线型 |

| 属性      | 作用       |
| --------- | ---------- |
| Linetypes | 返回线型集 |

------



* 块

| 方法        | 作用   |
| ----------- | ------ |
| CreateBlock | 创建块 |
| InsertBlock | 插入块 |

------



* 用户坐标系

| 方法          | 作用                    |
| ------------- | ----------------------- |
| CreateUCS     | 创建用户坐标系          |
| ActivateUCS   | 激活用户坐标系          |
| GetCurrentUCS | 获取当前用户坐标系      |
| ShowUCSIcon   | 显示/关闭用户坐标系图标 |

------



* 文字

| 方法              | 作用             |
| ----------------- | ---------------- |
| CreateTextStyle   | 创建文字样式     |
| ActivateTextStyle | 激活文字样式     |
| GetActiveFontInfo | 获取活动字体信息 |
| SetActiveFontFile | 设定活动字体文件 |
| AddText           | 创建单行文字     |
| AddMText          | 创建多行文字     |

------



* 尺寸与标注

| 方法             | 作用                     |
| ---------------- | ------------------------ |
| AddDimAligned    | 创建平行尺寸标注对象     |
| AddDimRotated    | 创建之地那个角度标注对象 |
| AddDimRadial     | 创建半径型尺寸标注对象   |
| AddDimDiametric  | 创建直径型尺寸标注对象   |
| AddDimAngular    | 创建角度型尺寸标注对象   |
| AddDimOrdinate   | 创建坐标型尺寸标注       |
| AddLeader        | 创建导线型标注           |
| CreateDimStyle   | 创建标注样式             |
| GetDimStyle      | 获取标注样式             |
| ActivateDimStyle | 激活标注样式             |

| 属性            | 作用                       |
| --------------- | -------------------------- |
| DimStyleNumbers | 返回标注样式数量           |
| DimStyleNames   | 返回标注样式名列表         |
| DimStyle0       | 返回index为0的标注样式对象 |
| DimStyles       | 返回标注样式集体           |
| ActiveDimStyle  | 返回当前标注样式           |

------



* Utility方法

Utility实际上就是与用户交互，比如用户输入字母，数字，做出选择等。



| 方法                | 作用               |
| ------------------- | ------------------ |
| GetString           | 获取字符串         |
| AngleFromXAxis      | 获取线与X轴的夹角  |
| GetPoint            | 获取空间一点的位置 |
| GetDistance         | 获取两点距离       |
| InitializeUserInput | 初始化用户输入选项 |
| GetKeyword          | 获取用户做出的选择 |
| GetEnity            | 点选实体           |
| GetReal             | 获取用户输入的实数 |
| GetInteger          | 获取用户输入的整数 |
| Prompt              | 给出提示           |

​		以上各种方法的详细用法，可以通过help命令查询，或者直接在源码中查询，不再赘述。

## 4.5 打印

​		打印目前没有直接纳入pycomcad中Autocad类方法，通过pycomcad来实现打印功能的用法详见博文：https://www.cnblogs.com/johnyang/p/14359725.html

# 5 与Pycomcad一起使用的第三方库

​		Python具备丰富而强大的第三方库，这也使快速开发出符合特定需求功能的Autocad二次开发程序成为可能。理论上，我们无法完全穷举出所有与Pycomcad一起使用的第三方库，因为特定需求本身就蕴含了无限可能，面对同一需求，实现的方法也不尽相同，使用的其他第三方库也不一样，下面列举出我自己在实际开发工作中常用到的其他第三方库，仅供参考。

| 第三方库  | 作用                                                         |
| --------- | ------------------------------------------------------------ |
| sys       | 刚需，将pycomcad所在路径添加到python搜索路径中，否则需要修改python的环境变量。 |
| os        | 在操作系统下进行的操作，如路径，文件的查询，添加等。         |
| shutil    | 更高级的文件操作库                                           |
| math      | 简单的数学运算                                               |
| numpy     | 向量化数值计算，避免了层层的for循环，在绘制不同比例的图形，非常有用 |
| pandas    | 与excel，csv交互，比如用excel上的数据来绘图，或者收集，储存图中的数据 |
| tkinter   | 图形界面化二次开发程序具有                                   |
| PyQt5     | 同tkinter，但比tkinter更为强大                               |
| itertools | 创建特定循环迭代器的函数集，比如数学上的排列，组合           |
| docx      | 与docx文件进行交互                                           |

# 6 从Autocad中调用二次开发程序

​		当我们通过pycomcad实现了某些自动化/自定义工作的功能后，如果需要频繁使用该程序，那么每次都直接运行脚本，显然有些繁琐，那么有没有办法可以从Autocad中直接调用写号的二次开发程序呢？

​		答案是有的，目前可以通过打包程序为exe文件（Pyinstaller的简单使用可参考我的博文：https://www.cnblogs.com/johnyang/p/10868863.html ），然后用lsp文件来调用打包的exe文件（不需要掌握lsp，很简单的一个语句），最后在Autocad中通过命令来调用该lsp文件就可以了。详见我的博文：https://www.cnblogs.com/johnyang/p/14415515.html

# 7 实战开发案例及不断升级...

​		 随着实际工作中遇到的开发需求越来越多，PycomCAD也在不断的升级中。如果你对该项目有任何的兴趣，可以clone它，尝试将它应用到实际工作中去，给本项目打个star。如果你发现有价值的功能需要被添加到PycomCAD中，可以pull request它，或者通过邮箱联系我：`1574891843@qq.com`。让我们一起携手，把PycomCAD打造的更为健壮，强大！

​		实战开发程序可参考 https://github.com/JohnYang1210/DesignWorkTask.

# 8 打赏及鸣谢

​		维护项目不易，如果您觉得该项目有帮到您，可以请博主喝杯咖啡~

  **微信二维码**

<img src="https://cdn.jsdelivr.net/gh/JohnYang1210/bloggitpic/img/20210227104234.jpg" alt="微信图片_20210227104225" style="zoom:50%;" />

​                                                   

------

  **支付宝二维码**                             

<img src="https://cdn.jsdelivr.net/gh/JohnYang1210/bloggitpic/img/20210227104423.jpg" alt="微信图片_20210227104410" style="zoom: 45%;" />

​                                                                                                   

**English version of Introduction:**

#  Introduction and development manual of PyComCAD

## 1.Overview

   In terms of the secondary development of Autocad in Engineering field, VB or Lisp may be choosen as the common and traditional programming language.However,Python shall play an important role as for this task with the power of easy-to-write and the elegance of conciseness, and Pycomcad exactly acts as an convenient way to get the API of Autocad.
    
   The base modules of Pycomcad are `win32com` and `pythoncom`,and win32com is responsible for getting the interface of Autocad including some constant values in the module level,pythoncom deals with the data type conversion.The methodology of Pycomcad is very easy and that is wrapping up calling functions of the multilayers to be single class methods or properties so that makes API function easier to memory and save your keystroke.
    
   When refering to the methods or properties of specific created entity in Autocad,It's better to look up `acadauto.chm` provided in this repository.

## 2.Base module installation

`pip install pywin32` will install both win32com and pythoncom.

## 3.Basic structure of Pycomcad

### 3.1 Module-level functions

These functions are used to convert data type.Details about data type convertion can be referred to https://www.cnblogs.com/johnyang/p/12617881.html .

### 3.2 Early-bind mode Or Lazy-bind mode

This blog(https://www.cnblogs.com/johnyang/p/12521301.html) written by myself may be consulted to learn about topics related to early-bind and lazy-bind mode.

By default,pycomcad is lazy-bind mode and that means pycomcad knows nothing about the method or property of specified entity,even the type of the entity itself.And actually, this has a huge impact on programming because we shall know clearly the type of entities in Autocad in order to do something different according to the type of selected entity.For example, as for EarlyBind mode,pycomcad will recognize win32com.client.constants.acRed,which is a constant value, while LazyBind mode will not recognize it.

Autocad object,assuming to `acad`, in Pycomcad has two properties to examin whether the module is earlybind or not and turn on earlybind mode if it is not,and they are `acad.IsEarlyBind` and `acad.TurnOnEarlyBind`.

Please note that,if there are multi-version Autocad on your PC, whether the Autocad object in pycomcad is EarlyBind or not will depend on the specific version of opened Autocad.So it's recommended to turn on all version's EarlyBind mode.

### 3.3 Major structure of module

* System variable
* File processing
* Precise drawing setting
* Entity creation
* Refer and select entity
* Layer
* Linetype
* Block 
* User-defined coordinate system
* Text
* Dimension and tolerance
* Utility object

Detailed information can be found in `pycomcad.py` and `acadauto.chm`.

### 4.Practical case and updating ...

Some actual application of pycomcad in my practical work can reffer to https://github.com/JohnYang1210/DesignWorkTask. With the increasing requirement encountered in daily work and for the integrity of module, pycomcad shall be evolving up to date constantly.
If you have any interest in this project,clone it,see the source and apply it to the real work.There are still so many function to add or update, once you find it,pull request it or contact with me through email:1574891843@qq.com, Let's work together to make Pycomcad more robust,integrated,concise and more powerful!

# 5 Donation

​		It's not easy for maintaining this project, and if you find it is useful , you can donate me a cup of caffee! 

  **WeChat QR  Code**

<img src="https://cdn.jsdelivr.net/gh/JohnYang1210/bloggitpic/img/20210227104234.jpg" alt="微信图片_20210227104225" style="zoom:50%;" />

​                                                                                                      

------

 

​    **Alipay QR Code**

<img src="https://cdn.jsdelivr.net/gh/JohnYang1210/bloggitpic/img/20210227104423.jpg" alt="微信图片_20210227104410" style="zoom: 45%;" />

​                                                                                                 