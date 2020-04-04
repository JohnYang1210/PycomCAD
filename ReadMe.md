
# PyComCAD tutorial and development manual

## 1.Overview

   In terms of the secondary development of Autocad in Engineering field, VB or Lisp may be choosen as the common and traditional programming language.However,Python shall play an important role as for this task with the power of easy-to-write and the elegance of conciseness, and Pycomcad exactly acts as an convenient way to get the API of Autocad.
    
   The base modules of Pycomcad are win32com and pythoncom,and win32com is responsible for get the interface of Autocad including some constant values in the module level,pythoncom deals with the data type conversion.The methodology of Pycomcad is very easy and that is wrapping up calling functions of the multilayers to be single class methods or properties so that makes API function easier to memory and save your keystroke.
    
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


```python

```
