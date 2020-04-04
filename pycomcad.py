"""
Based on COM technology, this module provides an easy way to employ the second-development of Autocad
"""
import pythoncom
import win32com.client
import math

def Apoint(x,y,z=0):
	"""
	Converts x,y,z into required float array as the arguments of coordinates of a point
	"""
	return win32com.client.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_R8,(x,y,z))  # | means the type combination

def ArrayTransform(x):
	"""
	x: any kind of array in python,such as ((1,2,3),(1,2,3),(1,2,3))
	"""
	return win32com.client.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_R8,x) 

def VtVertex(*args):
	"""
	Converts 2D coordinates of a serial points into the required float array
	"""
	return win32com.client.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_R8,args)

def VtObject(*obj):
	"""
	converts obj in python into required obj array
	"""
	return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, obj)
def VtFloat(list):
    """converts list in python into required float"""
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, list)
def VtInt(list):
    """converts list in python into required int"""
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, list)
def VtVariant(list):
    """converts list in python into required variant"""
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, list)
def AngleDtoR(degree):
	"""
	convert degree to rad
	"""
	rad=degree*math.pi/180
	return rad 
def AngleRtoD(rad):
	"""
	convert rad to degree
	"""
	degreee=180*rad/math.pi 
	return degreee

def FilterType(ftype):
	"""
	ftype shall be tuple.Refer to DXF in acad_aag.chm to learn about DXF group code.
	"""
	return win32com.client.VARIANT(pythoncom.VT_I2|pythoncom.VT_ARRAY,ftype)

def FilterData(fdata):
	"""
	fdata shall be tuple.Refer to DXF in acad_aag.chm to learn about DXF group code.
	"""
	return win32com.client.VARIANT(pythoncom.VT_VARIANT|pythoncom.VT_ARRAY,fdata)

class PycomError(Exception):
	def __init__(self,info):
		print(info)
class Autocad:
	def __init__(self):
		try:
			self.acad=win32com.client.Dispatch("Autocad.Application")
			self.acad.Visible=True 
		except:
			Autocad.__init__(self)
	"""
	Application
	"""
	@property
	def IsEarlyBind(self):
		"""
		Test whether it is
		:return:
		"""
		if 'IAcadApplication' in str(type(self.acad)):
			return True
		else:
			return False
	def TurnOnEarlyBind(self):
		import os,sys
		makepyPath=r'Lib\site-packages\win32com\client\makepy.py'
		ExePath=os.path.split(sys.executable)[0]
		MakePyPath=os.path.join(ExePath,makepyPath)
		os.execl(sys.executable,'python',MakePyPath)
	def AppPath(self):
		"""
		return Autocad Application path
		"""
		return self.acad.Path
	"""
	System variable
	"""
	def SetVariable(self,name,value):
		"""
		name:string
		"""
		self.acad.ActiveDocument.SetVariable(name,value)
	def GetVariable(self,name):
		return self.acad.ActiveDocument.GetVariable(name)
	"""
	File processing
	"""

	def OpenFile(self,path):
		"""
		open a dwg file in the path
		"""
		self.acad.Documents.Open(path)
	def CreateNewFile(self):
		"""
		Create a new dwg file,by default, the name is Drawing1.dwg
		"""
		self.acad.Documents.Add()
	def SaveFile(self):
		"""
		save file
		"""
		self.acad.ActiveDocument.Save()
	def SaveAsFile(self,path):
		"""
		save as file
		"""
		self.acad.ActiveDocument.SaveAs(path)
	def Close(self):
		"""
		Close current file
		:return:
		"""
		self.acad.ActiveDocument.Close()

	@property
	def OpenedFilenames(self):
		"""
		:return: list,all opened filenames
		"""
		names=[]
		for i in range(self.OpenedFilenumbers):
			names.append(self.acad.Documents.Item(i).Name)
		return names
	@property
	def OpenedFilenumbers(self):
		"""
		:return: the number of opened file
		"""
		return self.acad.Documents.Count

	def GetOpenedFile(self,file):
		"""
		Return already opened file whose index is index or name is name as
		the Current file
		:param file: int,or string,the number of index or the name of file targeted to be set as the current file
		"""
		if isinstance(file,str):
			index = self.OpenedFilenames.index(file)
		elif isinstance(file,int):
			index=file
		else:
			raise PycomError('Type of file in GetOpenedFile is wrong ')
		return self.acad.Documents.Item(index)

	def ActivateFile(self,file):
		"""
		Activate already opened file whose index is index or name is name as
		the Current file
		:param file: int,or str,the number of index or the name of file targeted to be set as the current file
		"""
		if isinstance(file,str):
			index = self.OpenedFilenames.index(name)
		elif isinstance(file,int):
			index=file
		else:
			raise PycomError('Type of file in ActivateFile() is wrong')
		self.acad.Documents.Item(index).Activate()

	def DeepClone(self,objects,Owner=None,IDPairs=win32com.client.VARIANT(pythoncom.VT_VARIANT, ())):
		"""
		Deep clone objects from current file to specified file's ModelSpace
		:param objects: objects needed to be deep cloned.Type: IAcadSelectionSet(selection sets),tuple of entity object
		:param Owner:specified opened file name,Type: string;Or the index of specified opened file name.
		:param ID:IDPairs.Default value has been set.
		:return:tuple of deep cloned object
		For example:
		>>from pycomcad import *
		>>acad=Autocad()
		>>te1=acad.AddCircle(Apoint(0,0,0),200)
		>>te2=acad.AddCircle(Apoint(100,100,0),200)
		>>acad.CreateNewFile()
		>>acad.ActivateFile(0)
		>>result=acad.DeepClone((te1,),1) # Deep Clone one object,notice the naunce between (te1,)and (te1),the latter one is int.
		>>result[0][0].Move(Apoint(0,0,0),Apoint(100,100,0))
		>>acad.CurrentFilename
		>>slt=acad.GetSelectionSets('slt1')
		>>slt.SelectOnScreen()
		>>result1=acad.DeepClone(slt,'Drawing2.dwg')
		"""
		if isinstance(objects,tuple):
			if not objects:
				raise PycomError('Objects in DeepClone() is empty tuple ')
			else:
				obj=VtObject(*objects)
		elif 'IAcadSelectionSet' in str(type(objects)):
			if objects.Count==0:
				raise PycomError('SelectionSets in DeepClone() is empty')
			else:
				obj=[]
				for i in range(objects.Count):
					obj.append(objects.Item(i))
				obj=VtObject(*obj)
		else:
			raise PycomError('Type of objects in DeepClone() is wrong')
		if not Owner:
			return self.acad.ActiveDocument.CopyObjects(obj)
		else:
			try:
				newOwner=self.GetOpenedFile(Owner).ModelSpace
			except:
				raise PycomError('File %s is not opened'% Owner)
			return self.acad.ActiveDocument.CopyObjects(obj,newOwner,IDPairs)

	@property
	def CurrentFilename(self):
		"""
		:return: str,the name of current file name
		"""
		return self.acad.ActiveDocument.Name

	@property
	def FilePath(self):
		"""
		return current file path
		"""
		return self.acad.ActiveDocument.Path
	@property
	def IsSaved(self):
		"""
		Specifies if the document has any unsaved changes
		:return:True: The document has no unsaved changes.
				False: The document has unsaved changes.
		"""
		return self.acad.ActiveDocument.Saved



	"""
	Zoom
	"""

	def ZoomExtents(self):
		self.acad.ZoomExtents()
	def ZoomAll(self):
		self.acad.ZoomAll()

	"""
	precise-drawing setting
	"""

	def GridOn(self,boolean):
		"""
		grid-on
		"""
		self.acad.ActiveDocument.ActiveViewport.GridOn=boolean
		self.acad.ActiveDocument.ActiveViewport=self.acad.ActiveDocument.ActiveViewport
	def SnapOn(self,boolean):
		"""
		snap-on
		"""
		self.acad.ActiveDocument.ActiveViewport.SnapOn=boolean
		self.acad.ActiveDocument.ActiveViewport.SnapOn=self.acad.ActiveDocument.ActiveViewport.SnapOn


	"""
	CAD entity Object drawing
	"""

	def AddPoint(self,apoint):
		"""
		apoint shall be got through Apoint function
		"""
		point=self.acad.ActiveDocument.ModelSpace.AddPoint(apoint)
		return point 

	def AddLine(self,startPoint,endPoint):
		"""
		The type of startPoint and endPoint are both Apoint
		"""
		line=self.acad.ActiveDocument.ModelSpace.AddLine(startPoint,endPoint)
		return line 

	def AddLwpline(self,*vertexCoord):
		"""
		LightWeight Poly line, this method requires the group of 2D vertex coordinates(that is x,y),
		This method is recommended to draw line
		"""
		lwpline=self.acad.ActiveDocument.ModelSpace.AddLightWeightPolyline(VtVertex(*vertexCoord))
		return lwpline

	def AddCircle(self,centerPnt,radius):
		"""
		add a circle, centerPnt's type is Apoint
		"""
		circle=self.acad.ActiveDocument.ModelSpace.AddCircle(centerPnt,radius)
		return circle

	def AddArc(self,centerPnt,radius,startAngle,endAngle):
		"""
		add an arc, startAngle and endAngle are both in the form of degree
		"""
		arc=self.acad.ActiveDocument.ModelSpace.AddArc(centerPnt,radius,AngleDtoR(startAngle),AngleDtoR(endAngle))
		return arc 

	def AddSpline(self,*fitPoints,startTan=None,endTan=None):
		"""
		fitpoints are array of 3D coordinates of points,such as (1,2,3,4,5,6)
		startTan is the starting vector which is the type of Apoint, and the same is endTan.
		"""
		spline=self.acad.ActiveDocument.ModelSpace.addSpline(VtVertex(*fitPoints),startTan,endTan)
		return spline

	def AddEllipse(self,centerPnt,majorAxis,radiusRatio):
		"""
		add ellipse, the type of majorAxis is Apoint
		"""
		ellipse=self.acad.ActiveDocument.ModelSpace.addEllipse(centerPnt,majorAxis,radiusRatio)
		return ellipse

	def AddHatch(self,patternType,patterName,associative,outLoopTuple,innerLoopTuple=None):
		"""
		The note of arguments can be seen as below:
		(1)patternType is the built-in integer constants which can be got by win32com.client.constants.X,here x can be 
		acHatchPatternTypeDefined(it means that using standard drawing file Acad.Pat to hatch, and the integer is 1),
		acHatchPatternTypeUserDefined(it means that using the current linetype to hatch,and the integer is 0),
		acHatchPatternTypeCustomDefined(it means that using user-defined drawing file .Pat to hatch,and the integer is 2)
		(2)patterName is a string specifying the hatch pattern name, such as "SOLID","ANSI31"
		(3)associative is boolean. If it is True, when the border is modified, the hatch pattern will adjust automatically 
		to keep in the modified border.
		(4)outLoop is a sequence of object,such as line,circle,etc. For example, outLoopTuple=(circle1,),or outLoopTuple=(line1,line2,
		line3).
		(5)innerLoop is the same with outLoop
		"""
		hatch=self.acad.ActiveDocument.ModelSpace.AddHatch(patternType,patterName,associative)
		out=VtObject(outLoopTuple)
		hatch.AppendOuterLoop(out)
		if innerLoopTuple:
			inn=VtObject(innerLoopTuple)
			hatch.AppendInnerLoop(inn)
		hatch.Evaluate()

	def AboutEntityObject(self):

		"""
		<This method is created only for the noting purpose>
		About editting autocad entity object:Users shall consult the acadauto.chm located in  C:\\Program Files
		\\Common Files\\Autodesk Shared\\acadauto.chm for exact supported property in terms of every kind of cad 
		entity.Some commen property and method has been summed up as below:
			(1)Commen Property:
				(a)object.color=X
				X:built-in contant, such as win32com.client.constants.acRed.Here,color is lowercase.
				(b)object.Layer=X
				X:string, the name of the layer
				(c)object.Linetype=X
				X:string,the name of the loaded linetype
				(d)object.LinetypeScale=X
				X:float,the linetype scale
				(e)object.Visible=X
				X:boolean,Determining whether the object is visible or invisible
				(f)object.EntityType
				read-only,returns an integer.
				(g)object.EntytyName
				read-only,returns a string
				(h)object.Handle
				read-only,returns a string
				(i)object.ObjectID
				read-only,returns a long integer
				(j)object.Lineweight=X
				X:built-in constants,(For example, win32com.client.constants.acLnWt030(0.3mm),acLnWt120 is
				1.2mm, and the scope of lineweight is 0~2.11mm),or acByLayer(the same with the layer
				where it lies),acByBlock,acBylwDefault.

			(2)Commen Method:
				(a)Copy
				RetVal=object.Copy
				RetVal: New created object
				object:Drawing entity,such as Arc,Line,LightweithPolyline,Spline,etc.

				(b)Offset
				RetVal=object.Offset(Distance)
				RetVal:New created object tuple
				Distance:Double,positive or negative
				object:Drawing entity,such as Arc,Line,LightweithPolyline,Spline,etc.

				(c)Mirror
				RetVal=object.Mirror(point1,point2)
				RetVal:mirror object
				point1,point2:end of mirror axis, Apoint type.
				object:Drawing entity,such as Arc,Line,LightweithPolyline,Spline,etc.

				(d)ArrayPolar
				RetVal=object.ArrayPolar(NumberOfObject,AngleToFill,CenterPoint)
				RetVal:New created object tuple
				NumberOfObject:integer,the number of array object(including object itself)
				AngleToFill:Double,rad angle, positive->anticlockwise,negative->clockwise
				CenterPoint:Double,Apoint type. The center of the array.
				object:Drawing entity,such as Arc,Line,LightweithPolyline,Spline,etc.

				(e)ArrayRectangular
				RetVal=object.ArrayRectangular(NumberOfRows,NumberOfColumns,NumberOfLevels,
				DistBetweenRows,DistBetweenColumns,DistBetweenLevels)
				RetVal:new created object tuple
				NumberOfRows,NumberOfColumns,NumberOfLevels:integer,the number of row,column,level,
				if it is the plain array that is performed, NumberOfLevels=1
				DistBetweenRows,DistBetweenColumns,DistBetweenLevels:Double,the distance between rows,
				columns,levels respectively.When NumberOfLevels=1,DistBetweenLevels is valid but still
				need to be passed
				object:Drawing entity,such as Arc,Line,LightweithPolyline,Spline,etc.

				(f)Move
				object.Move(point1,point2)
				object:Drawing entity,such as Arc,Line,LightweithPolyline,Spline,etc.
				point1,point2:Double,Apoint type. The moving vector shall be determined by the
				two points and point1 is the start point, point2 is the end point.

				(g)Rotate
				object.Rotate(BasePoint,RotationAngle)
				object:Drawing entity,such as Arc,Line,LightweithPolyline,Spline,etc.
				BasePoint:Double,Apoint type.The rotation basepoint.
				RotationAngle:Double,rad angle.

				(h)ScaleEntity
				object.ScaleEntity(BasePoint,ScaleFactor)
				object:Drawing entity,such as Arc,Line,LightweithPolyline,Spline,etc.
				BasePoint:Double,Apoint type.The scale basepoint.
				ScaleFactor:Double,Apoint type.

				(i)Erase
				object.Erase()
				object:Choosed set
				Delete all entity in the choosen scope

				(J)Delete
				object.Delete()
				object:specified entity, as for set object,such as modelSpace set and layerSet , this
				method is valid.

				(k)Update
				object.Update()
				update object after some kind of the objects' editing.

				(L)color
				object.color
				Here attention please, it is color,Not Color.(lowercase)

				(M)TransformBy
				object.TransformBy(transformationMatrix)
				object:Drawing entity,such as Arc,Line,LightweithPolyline,Spline,etc.
				transformationMatrix:4*4 Double array, need to be passed to ArrayTransform() method to be the required type 


			(3)Refer to Object

				(a)HandleToObject
				RetVal=object.HandleToObject(Handle)
				Retval:the entity object corresponding to Handle
				object:Document object
				Handle: the handle of entity object

				(b)ObjectIdToObject
				RetVal=object.ObjectIdToObject(ID)
				RetVal:the entity object corresponding to ID
				object:Document object
				ID: the identifier of object 
		"""
		pass

	"""
	Refer and select entity
	"""

	def GetEntityByItem(self,i):
		"""
		Refere to entity by its index location
		"""
		return self.acad.ActiveDocument.ModelSpace.Item(i)

	def GetSelectionSets(self,setname):
		"""
		setname:string, the name of selection set.
		There are 2 steps to select entity object:
		(1) create selection set 
		(2)Add entity into set
		Also note that: one set once has been created,
		it can never be created again, unless it is
		deleted.
		This method provides the first step.
		For example:
		>>>ft=[0, -4, 40, 8]  # define filter type
		>>>fd=['Circle', '>=', 5, '0'] #define filter data
		>>>ft=VtInt(ft) # data type convertion
		>>>fd=VtVariant(fd) #data type convertion
		>>>slt=acad.GetSelectionSets('slt') # Create selectionset object
		>>>slt.SelectOnScreen(ft,fd) # select on screen
		>>>slt.Erase() # Erase selected entity
		>>>slt.Delete() # Delete selectionsets object
     
		"""
		return self.acad.ActiveDocument.SelectionSets.Add(setname)

	"""
	There are 5 methods to add entity into selection set:

	(1)object.AddItems(Items)
	object:selection set
	Items:Variant tuple. For example, Items=VtObject((c1,c2)),where c1,c2 
	is the object being ready to join in selection set

	(2)object.Select(Mode[,Point1][,Point2][,FilterType][,FilterData])
	object:selection set
	Mode=win32com.client.constants.X
	X is as below:
	acSelectionSetWindow,acSelectionSetPrevious,acSelectionSetLast,
	acSelectionSetAll
	Point1,Point2: 2 diagonal points defining a window
	FilterType,FilterData: DXF group code; filter type. 

	(3)object.SelectAtPoint(Point[,FilterType][,FilterData])
	object:selection set
	Point:Given point

	(4)object.SelectByPolygon(Mode,PointsLists[,FilterType][,FilterData])
	object:selection set
	Mode=win32com.client.constants.X
	X is as below:
	acSelectionSetFence,acSelectionSetWindowPolygon,acSelectionSetCrossingPolygon
	PointsLists:a serial points(3D) defining polygon
	FilterType,FilterData: DXF group code; filter type.

	(5)object.SelectOnScreen()
	object:selection set
	FilterType,FilterData: DXF group code; filter type.
	"""

	"""
	Filter Mechanism:
	DXF                  filter type
	0              entity ,such as Line,Circle,Ac,etc.
	2              name of object (string)
	5              entity handle
	8              layer
	60             visible of entity
	62             color integer,0->BYBLOCK,256->BYLAYER,negative->closed layer
	67             ignored or 0->ModelSpace,1->PaperSpace

	DXF shall be passed into FilterType() in the form of tuple to be the required type,
	while filter type shall be passed into FilterData() in the form of tuple to be the 
	required type.
	"""


	"""
	Deletion of selection set:
	(1)Clear:clear the selection set, the selection set still exists and the member entities still
	exist but they no longer belong to this selection set.

	(2)RemoveItems:the removed member entities still exist, but they no longer belong to this selection
	set.

	(3)Erase:delete all the member entities and the selection set itsel still exists.

	(4)Delete:delete the selection set itself, but the member entities still exist.

	"""

	"""
	Layer
	"""
	def CreateLayer(self,layername):
		"""
		create new layer named layername(string)
		"""
		return self.acad.ActiveDocument.Layers.Add(layername)


	def ActivateLayer(self,layer):
		"""
		Activate layer
		layer:str or int, the index or the name of the being activated layer
		"""
		self.acad.ActiveDocument.ActiveLayer=self.GetLayer(layer)

	@property
	def LayerNumbers(self):
		"""
		return the number of layers in the active document
		"""
		return self.acad.ActiveDocument.Layers.Count

	@property
	def LayerNames(self):
		"""
		:return a list containing all layer names
		"""
		a=[]
		for i in range(self.LayerNumbers):
			a.append(self.acad.ActiveDocument.Layers.Item(i).Name)
		return a
	def GetLayer(self,layer):
		"""
		get an indexed layer
		layer:int or string, the index or the name of layer which exists already.
		"""
		if isinstance(layer,str):
			index=self.LayerNames.index(layer)
		elif isinstance(layer,int):
			index=layer
		else:
			raise PycomError('Type of layer in GetLayer() is wrong')
		return self.acad.ActiveDocument.Layers.Item(index)
	@property 
	def Layers(self):
		"""
		:return layer set object
		"""
		return self.acad.ActiveDocument.Layers
	@property
	def ActiveLayer(self):
		"""
		:return: ActiveLayer object
		"""
		return self.acad.ActiveDocument.ActiveLayer
	"""
	The state change and deletion of layer:
	(1)Obj.LayerOn=True/False
	Obj:Layer object
	closed or not,if it is  closed, new entity object can be created on layer,while it cannot be seen.
	(2)Obj.Freeze=True/False
	Obj:Layer object
	if freezed,the layer can neighter be shown or created entities on it.
	(3)Obj.Lock=True/False
	Obj:Layer object
	The entity on a locked layer can be shown, if the locked layer is activated, new entity can be 
	created there, but the entities cannot be edited or deleted.
	(4)Obj.Delete
	Obj:Layer object
	Delete any layer, except for cunnrent layer and 0 layer(default layer).

	The property of layer:
	(1)Obj.color=X
	X:built-in contant, such as win32com.client.constants.acRed
	(2)Obj.Linetype=X
	X:string, the name of loaded linetype
	(3)Obj.Name
	"""
	"""
	Linetype
	"""
	def LoadLinetype(self,typename,filename='acad.lin'):
		"""
		typename:string, the name of type needed to be load.such as 'dashed','center'
		filename:string, the name of the file the linetype is in.'acad.lin','acadiso.lin'
		"""
		self.acad.ActiveDocument.Linetypes.Load(typename,filename)

	def ActivateLinetype(self,typename):
		"""
		typename:string, ensure the typename has been loaded successfully.
		"""
		try:
			self.acad.ActiveDocument.ActivateLinetype=self.acad.ActiveDocument.Linetypes.Item(typename)
		except:
			print('The typename has not been loaded')

	def ShowLineweight(self,TrueorFalse):
		"""
		TrueorFalse:Boolean, determining whether the lineweight be shown or not
		"""
		self.acad.ActiveDocument.Preferences.LineWeightDisplay(TrueorFalse)
	@property
	def Linetypes(self):
		"""
		return linetype set
		"""
		return self.acad.ActiveDocument.Linetypes 
		
	

	"""
	Block
	There are 3 steps as for the creation and reference about Block.
	(1)create a block, see  the following method CreateBlcok
	(2)The created blcok adds enity;
	Obj.AddX
	X can be entity object,text object,etc.
	(3)insert block, see the fowllowing method InsertBlcok

	Block Explode
	Obj.Explode()
	Obj:Reference Block object
	This method returns a tuple containing the exploded object

	Block attribute object
	Retval=blockObj.AddAttribute(height,mode,prompt,insertPoint,tag,value)
	blockObj:Block reference object
	Retval:Attribute object
	height:Double float, the height of text
	Mode:built-in constants,win32com.client.constants.X,and X is as the following
		acAttributeModeInvisible:the attribute value is invisible
		acAttributeModeConstant:constant attribute, cannot be editted
		acAttributeModeVerify:when inserting block, prompt users to ensure the attribute value
		acAttributeModePreset:when inserting block, use default attribute value, cannot be editted
		These constants can be used as a combination
	
	GetAttribute method
		To access an attribute reference of an inserted block, use the GetAttributes method. 
		This method returns an array of all attribute references attached to the inserted block. 
	Retval=obj.GetAttributes()
	obj:Block reference object
	Retval:Block attribute object tuple
	Retval's 2 main property:(1)TagString(2)TextString
	Note:since Retval is a tuple, we may use len() method to get the number of the member in it
	"""

	def CreateBlock(self,insertPnt,blockName):
		"""
		insertPnt:Apoint type,the insertion base point
		blockName:string,the name of the new-created block
		"""
		return self.acad.ActiveDocument.Blocks.Add(insertPnt,blockName)

	def InsertBlock(self,insertPnt,blockName,Xscale=1,Yscale=1,Zscale=1,Rotation=0):
		"""
		insertPnt:Apoint type,the insert point in the process of block insertion.
		blockName:string, the inserted block name which has been created
		"""
		return self.acad.ActiveDocument.ModelSpace.InsertBlock(insertPnt,blockName,Xscale,Yscale,Zscale,Rotation)

	"""
	User-defined coordinate system
	Normally, users perform drawing work in WCS(world coordinate system).However,in some case, it is easy
	to draw in UCS(user coordinate system). In UCS, it's necessary to use coordinate transform, and the steps
	are as follow:
	(1)Create entity in WCS directly
	(2)Create UCS and get transform matrix of UCS by method GetUCSMatrix (here, also need array type conversion
	 by ArrayTransform method)
	(3)Transform the entity created in WCS to UCS through method TransformBy
	Also attention that after the transform perform , it's better to set the previous coordinate system.

	TransMatrix=ucsObj.GetUCSMatrix()
	TransMatrix:4*4 Double array, need to be passed to ArrayTransform() method to be the required type 
	ucsObj:UCS object

	TransformBy
	object.TransformBy(transformationMatrix)
	object:Drawing entity,such as Arc,Line,LightweithPolyline,Spline,etc.
	transformationMatrix:4*4 Double array, need to be passed to ArrayTransform() method to be the required type 
	"""

	def CreateUCS(self,origin,xAxisPnt,yAxisPnt,csName):
		"""
		origin:Apoint type,origin point of the new CS
		xAxisPnt:Apoint type,one point directing the positive direction of x axis of the new CS
		yAxisPnt:Apoint type,one point directing the positive direction of y axis of the new CS
		csName:string,the name of the new CS
		"""
		return self.acad.ActiveDocument.UserCoordinateSystems.Add(origin,xAxisPnt,yAxisPnt,csName)

	def ActivateUCS(self,ucsObj):
		"""
		ucsObj: CS object
		"""
		self.acad.ActiveDocument.ActiveUCS=ucsObj

	def GetCurrentUCS(self):
		"""
		Before activate the new created UCS, it's better to get the current UCS in case of activating
		it after tasks in the activated new created UCS.
		"""
		if self.acad.ActiveDocument.GetVariable('ucsname')=='':
			origin=self.acad.ActiveDocument.GetVariable('ucsorg')
			origin=ArrayTransform(origin)
			xAxisPnt=self.acad.ActiveDocument.Utility.TranslateCoordinates(ArrayTransform(self.acad.ActiveDocument.GetVariable('ucsxdir')),
				win32com.client.constants.acUCS,win32com.client.constants.acWorld,0)
			xAxisPnt=ArrayTransform(xAxisPnt)
			yAxisPnt=self.acad.ActiveDocument.Utility.TranslateCoordinates(ArrayTransform(self.acad.ActiveDocument.GetVariable('ucsydir')),
				win32com.client.constants.acUCS,win32com.client.constants.acWorld,0)
			yAxisPnt=ArrayTransform(yAxisPnt)
			currCS=self.acad.ActiveDocument.UserCoordinateSystems.Add(origin,xAxisPnt,yAxisPnt,'currentUCS')
		else:
			currCS=self.acad.ActiveDocument.ActiveUCS
		return currCS

	def ShowUCSIcon(self,booleanOfUCSIcon,booleanOfUCSatOrigin):
		"""
		show UCS Icon
		booleanOfUCSIcon:boolean,Specifies if the UCS icon is on
		booleanOfUCSatOrigin:boolean,Specifies if the UCS icon is displayed at the origin
		"""
		self.acad.ActiveDocument.ActiveViewport.UCSIconOn=booleanOfUCSIcon
		self.acad.ActiveDocument.ActiveViewport.UCSIconAtOrigin=booleanOfUCSatOrigin

	"""
	Text

	Text Style Object
		(1)SetFont method
		object.SetFont(TypeFace,Bold,Italic,CharSet,PitchandFamily)
		Function->Set the font for created text style object
		object:text style object
		TypeFace:string, font name, such as "宋体"
		Bold:boolean,if True, bold, if False, normal
		Italic:boolean, if True, italic,if False,normal
		CharSet: long integer, defining font character set, the constants's meaning is as below
			Constant 			Meaning
			0					ANSI character set
			1 					Default character set
			2 					Symbol set
			128 				Japanese character set
			255					OEM character set 
		PitchandFamily: consists of 2 part:(a)Pitch,defining character's pitch(b)Family,defining character'stroke
			Pitch:
				Constant 					Meanning
				0 							Default value
				1 							Fixed value
				2 							variable value
			Family:
				Conatant 					Meanning
				0 							No consideration of stroke form
				16							Variable stroke width,with serif
				32 							Variable stroke width,without serif
				48 							Fixed stroke width,with or without serif
				64 							Grass writting
				80 							Old English stroke
		(2) FontFile property
		obj.fontFile=path
		obj:textstyle object
		set the given textstyle's font file by the path of character file,
		for example, path=self.acad.Path+r'\tssdeng.shx'

		(3)BigFontFile property
		obj.BigFontFile=path
		obj:textstyle object
		This property is similar to the FontFile property, except that it is used to specify 
		an Asian-language Big Font file. The only valid file type is SHX

	"""

	def CreateTextStyle(self,textStyleName):
		"""
		testStyleName:string,the name of new created text style
		"""
		return self.acad.ActiveDocument.TextStyles.Add(textStyleName)

	def ActivateTextStyle(self,textStyleObj):
		"""
		Activate the created textstyle object
		textStyleObj:textStyle object
		"""
		self.acad.ActiveDocument.ActiveTextStyle=textStyleObj

	def GetActiveFontInfo(self):
		"""
		return a tuple (typeFace,Bold,Italic,charSet,PitchandFamily) of the active textstyle object
		"""
		return self.acad.ActiveDocument.ActiveTextStyle.GetFont()

	def SetActiveFontFile(self,path):
		"""
		set the active textstyle's font file by the path of character file,
		for example, path=self.acad.Path+r'\tssdeng.shx'
		"""
		self.acad.ActiveDocument.ActiveTextStyle.fontFile=path 
	def SetActiveBigFontFile(self,path):
		"""
		This property is similar to the FontFile property, except that it is used to specify 
		an Asian-language Big Font file. The only valid file type is SHX
		"""
		self.acad.ActiveDocument.ActiveTextStyle.BigFontFile=path

	"""
	Single Text

	Formatted text
		(1)Alignment
			object.Alignment=win32com.client.constants.X
			[object.TextAlignmentPoint=pnt1]
			[object.InsertionPoint=pnt2]
			object:single text object
			X:acAlignmentLeft 
			acAlignmentCenter 
			acAlignmentRight 
			acAlignmentAligned 
			acAlignmentMiddle 
			acAlignmentFit 
			acAlignmentTopLeft 
			acAlignmentTopCenter 
			acAlignmentTopRight 
			acAlignmentMiddleLeft 
			acAlignmentMiddleCenter 
			acAlignmentMiddleRight 
			acAlignmentBottomLeft 
			acAlignmentBottomCenter 
			acAlignmentBottomRight
		Note that:Alignment property has to be set before TextAlignmentPoint or InsertionPoint property be set!
		Text aligned to acAlignmentLeft uses the InsertionPoint property to position the text. Text aligned to 
		acAlignmentAligned or acAlignmentFit uses both the InsertionPoint and TextAlignmentPoint properties to
		position the text. Text aligned to any other position uses the TextAlignmentPoint property to position the text.

		(2)InsertionPoint
			object.InsertionPoint=pnt
			pnt:Apoint type
		Note:This property is read-only except for text whose Alignment property is set to acAlignmentLeft, 
		acAlignmentAligned, or acAlignmentFit. To position text whose justification is other than left, aligned,
		or fit, use the TextAlignmentPoint property.

		(3)ObliqueAngle
			object.ObliqueAngle=rad
			rad:Double,rad angle
			The angle in radians within the range of -85 to +85 degrees. A positive angle denotes a lean to the right; 
			a negative value will have 2*PI added to it to convert it to its positive equivalent. 

		(4)Rotation
			object.Rotation=rad
			rad:Double,The rotation angle in radians. 

		(5)TextAlignmentPoint
		objcet.TextAlignmentPoint=pnt
		pnt：Apoint type
		Specifies the alignment point for text and attributes;Note that:Alignment property has to be set before 
		TextAlignmentPoint or InsertionPoint property be set!Text aligned to acAlignmentLeft uses the InsertionPoint 
		property to position the text.

		(6)TextGenerationFlag
		object.TextGenerationFlat=win32com.client.constants.x
		X:acTextFlagBackward,acTextFlagUpsideDown
		Specifies the attribute text generation flag,To specify both flags, add them together,
		that is acTextFlagBackward+acTextFlagUpsideDown

		(7)TextString
		object.TextString
		This method returns the text string of single text object

		(8)commen editing method:
		ArrayPolar,ArrayRectangular,Copy,Delete,Mirror,Move,Rotate.
	"""
	def AddText(self,textString,insertPnt,height):
		"""
		add single text
		textString:string,the inserted single text
		insertPnt:Apoint type,insert point
		height:the text height
		"""
		return self.acad.ActiveDocument.ModelSpace.AddText(textString,insertPnt,height)
	"""
	MutiText
	"""
	def AddMText(self,textString,insertPnt,width):
		"""
		Creates an MText entity in a rectangle defined by the insertion point and width of the bounding box.
		textString:string
		insertPnt:Apoint type
		width:float,The width of the MText bounding box
		"""
		return self.acad.ActiveDocument.ModelSpace.AddMText(insertPnt,width,textString)

	"""
	Dimension and Tolerance

	Common property of dim object
		(1)obj.DecimalSeparator=X
		X:string,such as '.',can be any string.

		(2)obj.ArrowheadSize=X
		X:Double,The size of the arrowhead must be specified as a positive real >= 0.0,The initial value for this property is 0.1800.

		(3)obj.DimensionLineColor=X
		X:Use a color index number from 0 to 256, or bilt-in constants

		(4)obj.DimLineInside=X
		X:Boolean, default is False. Specifies the display of dimension lines inside the extension lines . Dimension line is the line below
		the dimenion text and extension lines are a pair of lines pointing to the limit point of a dimension.

		(5)obj.Fit=win32com.client.constants.X
		Specifies the placement of text and arrowheads inside or outside extension lines, based on the available space between the extension lines
		X:acTextAndArrows,acArrowsOnly,acTextOnly,acBestFit

		(6)obj.Measurement
		Read-only,returns the actural dimension value.

		(7)obj.TextColor
		the text color

		(8)obj.TextHeight
		the text height

		(9)obj.TextOverride=X
		X:string. ''represents the actural measurement.'<>'represents the actural measurment value,such as '<>mm'

		(10)obj.Arrowhead1Type=win32com.client.constants.X
		obj.Arrowhead2Type=win32com.client.constants.X
		X:
			acArrowDefault,acArrowDot,acArrowDotSmall,acArrowDotBlank,acArrowOpen,acArrowOblique,acArrowArchTick,etc.

		(11)obj.TextPosition=X
		X:Apoint type. the position of text.

		(12)obj.TextPrefix=X
		X:string

		(13)obj.TextSuffix=X
		X:string
		(14)obj.UnitsFormat=win32com.client.constants.X
		Specifies the unit format for all dimensions except angular
		X:
			acDimLScientific,acDimLDecimal,acDimLEngineering,acDimLArchitectural,acDimLFractional
		The initial value for this property is acDimLDecimal.If this property is set to acDimLDecimal, 
		the format specified by the DecimalSeparator and PrimaryUnitsPrecision properties will be used to format the decimal value

		(15)obj.PrimaryUnitsPrecision=win32com.client.constants.X
		Specifies the number of decimal places displayed for the primary units of a dimension or tolerance
		X:
			acDimPrecisionZero: 0
			acDimPrecisionOne: 0.0
			acDimPrecisionTwo: 0.00
			acDimPrecisionThree: 0.000
			acDimPrecisionFour: 0.0000 
			acDimPrecisionFive: 0.00000
			acDimPrecisionSix: 0.000000
			acDimPrecisionSeven: 0.0000000
			acDimPrecisionEight: 0.00000000 

		(16)obj.VerticalTextPosition=win32com.client.constants.X
		X:
			acAbove,acOutside,acVertCentered,acJI

		(17)obj.TextOutsideAlign=X
		obj:
		X:Boolean,Specifies the position of dimension text outside the extension lines for all dimension types except ordinate
		True: Align the text horizontally
		False: Align the text with the dimension line

		(18)obj.CenterType=win32com.client.constants.X
		Specifies the type of center mark for radial and diameter dimensions
		obj: DimDiametric, DimRadial, DimRadialLarge 
			X:
			acCenterMark 
			acCenterLine 
			acCenterNone
		Note:The center mark is visible only if you place the dimension line outside the circle or arc.

		(19) obj.CenterMarkSize=X
		Specifies the size of the center mark for radial and diameter dimensions.
		X:Double,A positive real number specifying the size of the center mark or lines
		Note:The initial value for this property is 0.0900. This property is not available if the CenterType property is set to acCenterNone.

		(20)obj.ForceLineInside=X
		Specifies whether a dimension line is drawn between the extension lines even when the text is placed outside the extension lines
		X:Boolean
		True: Draw dimension lines between the measured points when arrowheads are placed outside the measured points. 
		False: Do not draw dimension lines between the measured points when arrowheads are placed outside the measured points

		(21)obj.StyleName=X
		X:string
		Specifies the name of the style used with the object





		




	"""
	def AddDimAligned(self,extPnt1,extPnt2,textPosition):
		"""
		Creates an aligned dimension object
		extPnt1:Apoint type,the 3D WCS coordinates specifying the first endpoint of the extension line
		extPnt2:Apoint type,the 3D WCS coordinates specifying the second endpoint of the extension line
		textPosition:Apoint type,the 3D WCS coordinates specifying the text position
		"""
		return self.acad.ActiveDocument.ModelSpace.AddDimAligned(extPnt1,extPnt2,textPosition)
	def AddDimRotated(self,xlPnt1,xlPnt2,dimLineLocation,rotAngle):
		"""
		Creates a rotated linear dimension
		xlPnt1:Apoint type,the 3D WCS coordinates specifying the first endpoint of the extension line
		xlPnt2:Apoint type,the 3D WCS coordinates specifying the first endpoint of the extension line
		rotAngle:Double,The angle, in radians, of rotation displaying the linear dimension
		"""
		return self.acad.ActiveDocument.ModelSpace.AddDimRotated(xlPnt1,xlPnt2,dimLineLocation,rotAngle)
	def AddDimRadial(self,center,chordPnt,leaderLength):
		"""
		Creates a radial dimension for the selected object at the given location
		center:Apoint tyoe
		chordPnt:Apoint type,The 3D WCS coordinates specifying the point on the circle or arc to attach the leader line
		leaderLength:double,The positive value representing the length from the ChordPoint to the annotation text or dogleg
		"""
		return self.acad.ActiveDocument.ModelSpace.AddDimRadial(center,chordPnt,leaderLength)
	def AddDimDiametric(self,chordPnt,farChordPnt,leaderLength):
		"""
		Creates a diametric dimension for a circle or arc given the two points on the diameter and the length of the leader line
		chordPnt:Apoint type,The 3D WCS coordinates specifying the first diameter point on the circle or arc
		farChordPnt:Apoint type,The 3D WCS coordinates specifying the second diameter point on the circle or arc
		leaderLength:The positive value representing the length from the ChordPoint to the annotation text or dogleg, when it is 0,
		using obj.Fit=win32com.client.constants.acTextAndArrows can make the arrow and text inside the circle

		"""
		return self.acad.ActiveDocument.ModelSpace.AddDimDiametric(chordPnt,farChordPnt,leaderLength)
	def AddDimAngular(self,vertex,firstPnt,secondPnt,textPnt):
		"""
		Creates an angular dimension for an arc, two lines, or a circle
		vertex,Apoint type,The 3D WCS coordinates specifying the center of the circle or arc, or the common vertex between the two dimensioned lines
		firstPnt,Apoint type,The 3D WCS coordinates specifying the point through which the first extension line passes
		secondPnt,Apoint type,The 3D WCS coordinates specifying the point through which the second extension line passes
		textPnt,Apoint type,The 3D WCS coordinates specifying the point at which the dimension text is to be displayed
		"""
		return self.acad.ActiveDocument.ModelSpace.AddDimAngular(vertex,firstPnt,secondPnt,textPnt)
	def AddDimOrdinate(self,definitionPnt,leaderPnt,axis):
		"""
		Creates an ordinate dimension given the definition point and the leader endpoint
		definitionPnt,Apoint type,The 3D WCS coordinates specifying the point to be dimensioned
		leaderPnt,Apoint type,The 3D WCS coordinates specifying the endpoint of the leader. This will be the location at which the dimension text is displayed
		axis,Boolean,True: Creates an ordinate dimension displaying the X axis value;False: Creates an ordinate dimension displaying the Y axis value
		"""
		return self.acad.ActiveDocument.ModelSpace.AddDimOrdinate(definitionPnt,leaderPnt,axis)
	def AddLeader(self,*pntArray,annotation=None,type=None):
		"""
		Creates a leader line based on the provided coordinates or adds a new leader cluster to the MLeader object
		pntArray,The array of 3D WCS coordinates,such as (1,2,3,4,5,6), specifying the leader. You must provide at least two points to define the leader. The third point is optional
		annotation,BlockReference,MText,Tolerance type.The object that should be attached to the leader. The value can also be NULL to not attach an 
		Type:built-in contants, win32com.client.constants.X,X is as the following:
		acLineNoArrow 
		acLineWithArrow 
		acSplineNoArrow 
		acSplineWithArrow
		>>>ann=acad.AddMText('demo',Apoint(30,30,0),2)
		>>>import win32com.client
		>>>acad.AddLeader(0,0,0,30,30,0,annotation=a,type=win32com.client.constants.acLineWithArrow)
		"""
		return self.acad.ActiveDocument.ModelSpace.AddLeader(VtVertex(*pntArray),annotation,type)

	"""
	Dimension style object

	(1)obj.CopyFrom(X)
	X:self.DimStyle0, self.ActiveDimStyle,and other dimension style object

	"""
	def AddDimStyle(self,name):
		"""
		creat a new dimension style named name.
		name:string
		"""
		return self.acad.ActiveDocument.DimStyles.Add(name)
	@property 
	def DimStyle0(self):
		"""
		return created dimension style object whose index is 0 in modelspace,

		"""
		return self.acad.ActiveDocument.ModelSpace(0)
	@property 
	def DimStyles(self):
		"""
		return all dimstyles
		"""
		return self.acad.ActiveDocument.DimStyles
	@property 
	def ActiveDimStyle(self):
		"""
		return a dim style set by system variable
		"""
		return self.acad.ActiveDocument

	"""
	Utility object
	"""
	def AngleFromXAxis(self,pnt1,pnt2):
		"""
		Gets the angle of a line from the X axis
		pnt1:Apoint type,The start point of the line;
		pnt2:Apoint type,The endpoint of the line.
		"""
		return self.acad.ActiveDocument.Utility.AngleFromXAxis(pnt1,pnt2)
	def GetAngle(self,basePnt=Apoint(0,0,0),prompt=''):
		"""
		Gets the angle specified
		"""
		return self.acad.ActiveDocument.Utility.GetAngle(basePnt,prompt)
	def GetPoint(self):
		return self.acad.ActiveDocument.Utility.GetPoint()
	def GetDistance(self,pnt='',prompt=''):
		"""
		Gets the point selected in AutoCAD
		pnt:Apoint type,optional,The Point parameter specifies a relative base point in the WCS
		"""
		if not pnt:
			return self.acad.ActiveDocument.Utility.GetDistance(ArrayTransform(self.GetPoint()),prompt)
		else:
			return self.self.acad.ActiveDocument.Utility.GetDistance(pnt,prompt)
	def InitializeUserInput(self,bits,keywords):
		"""
		Before using GetKeyword method,this method has to be used to limit the user-input forms , and this method
		can also used with GetAngle,GetCorner,GetDistance,GetInteger,GetOrientation,GetPoint,GetReal, and cannot be
		used with GetString.Unless it is set again,or it will control the type of input forever.
		bits:integer
			1: Disallows NULL input. This prevents the user from responding to the request by entering only [Return] or a space. 
			2: Disallows input of zero (0). This prevents the user from responding to the request by entering 0. 
			4: Disallows negative values. This prevents the user from responding to the request by entering a negative value. 
			8: Does not check drawing limits, even if the LIMCHECK system variable is on. This enables the user to enter a point outside the current drawing limits. This condition applies to the next user-input function even if the AutoCAD LIMCHECK system variable is currently set. 
			16: Not currently used. 
			32: Uses dashed lines when drawing rubber-band lines or boxes. This causes the rubber-band line or box that AutoCAD displays to be dashed instead of solid, for those methods that let the user specify a point by selecting a location on the graphics screen. (Some display drivers use a distinctive color instead of dashed lines.) If the AutoCAD POPUPS system variable is 0, AutoCAD ignores this bit. 
			64: Ignores Z coordinate of 3D points (GetDistance method only). This option ignores the Z coordinate of 3D points returned by the GetDistance method, so an application can ensure this function returns a 2D distance. 
			128: Allows arbitrary input—whatever the user types. 
		keywords:strings,such as 'width length height'

		"""
		self.acad.ActiveDocument.Utility.InitializeUserInput(bits,keywords)
	def GetKeyword(self,prompt=''):
		"""
		Before using GetKeyword method,this method has to be used to limit the user-input forms
		Gets a keyword string from the user
		"""
		return self.acad.ActiveDocument.Utility.GetKeyword(prompt)
	def GetEntity(self):
		"""
		Return a tuple containing the picked object and the coordinate of  picked point
		"""
		return self.acad.ActiveDocument.Utility.GetEntity()
	def GetReal(self,prompt=''):
		"""
		Gets a real (double) value from the user.
		"""
		return self.acad.ActiveDocument.Utility.GetReal(prompt)

# 	def GetSelectionSets(self,string):
# 		"""
# 				:param string: name of selection set
# 				:return: SelectionSets object default name is '0'
# 				For example:
# 				>>>ft=[0, -4, 40, 8]  # define filter type
# 				>>>fd=['Circle', '>=', 5, '0'] #define filter data
# 				>>>ft=VtInt(ft) # data type convertion
# 				>>>fd=VtVariant(fd) #data type convertion
# 				>>>slt=acad.GetSelectionSets() # Create selectionset object
# 				>>>slt.SelectOnScreen(ft,fd) # select on screen
# 				>>>slt.Erase() # Erase selected entity
# 				>>>slt.Delete() # Delete selectionsets object
# 				"""
# 		return self.acad.ActiveDocument.SelectionSets.Add (string)

if __name__=='__main__':
	table={'dimclrd':62,'dimdlI':0,'dimclre':62,
	   'dimexe':2,'dimexo':3,
	   'dimfxlon':1,
	   'dimfxl':3,'dimblk1':'_archtick',
	   'dimldrblk':'_dot',
	   'dimcen':2.5,'dimclrt':62,'dimtxt':3,'dimtix':1,
	   'dimdsep':'.','dimlfac':50}
	acad=Autocad()
	p1=Apoint(0,0,0)
	c1=acad.AddCircle(p1,10)
	c2=acad.AddCircle(p1,5)
	acad.ZoomExtents()
	acad.AddHatch(1,'SOLID',True,(c1,),(c2,))





