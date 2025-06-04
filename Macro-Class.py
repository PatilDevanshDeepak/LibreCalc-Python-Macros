import uno

class WorkBook:
	def __init__(self):
		self.doc = XSCRIPTCONTEXT.getDocument()
		self.active = self.doc.CurrentController
		self.ActiveSheet = self.getActiveSheet()
		self.ActiveCell = self.getActiveSelection()
		
	# Macro Functions a
	def Select(self, oCell):
		self.active.select(oCell)
		self.ActiveCell = oCell.getCellByPosition(0 , 0)
	
	def Cell(self, row, col):
		return self.ActiveSheet.getCellByPosition(col-1, row-1)
		
	def Range(self, rangeName=None , fromCell=None , toCell=None):
		if self.isRangeNameGiven(isRowIndex=fromCell , isColumnIndex=toCell , isRangeName=rangeName):
			return self.ActiveSheet.getCellRangeByName(rangeName)
		else:
			return self.ActiveSheet.getCellRangeByName(fromCell.AbsoluteName+":"+toCell.AbsoluteName)

	def getStartColumn(self , oCell):
		return oCell.RangeAddress.StartColumn

	def getStartRow(self , oCell):
		return oCell.RangeAddress.StartRow

	def getEndColumn(self , oCell):
		return self.ActiveSheet.Columns.Count

	def getEndRow(self , oCell):
		return self.ActiveSheet.Rows.Count
		
	def Offset(self, rowIndex, colIndex):
		maxColumns = self.ActiveSheet.Columns.Count - 1
		maxRows = self.ActiveSheet.Rows.Count - 1
		newRow = self.ActiveCell.RangeAddress.StartRow + rowIndex
		newCol = self.ActiveCell.RangeAddress.StartColumn + colIndex
		newRow = max(0, min(newRow, maxRows))
		newCol = max(0, min(newCol, maxColumns))
		return self.ActiveSheet.getCellByPosition(newCol, newRow)
		
	def Row(self, oCell):
		return 	oCell.RangeAddress.StartRow + 1
	
	def Column(self, oCell):
		return oCell.RangeAddress.StartColumn + 1
	
	def MsgBox(self, message="none", title="Information", msgBoxType=0):					# Issues with this function
		self.active.getFrame().getComponentWindow().showMessageBox(message, title, msgBoxType)

	def clcToRight(self):				# This function is like xlToRight in VBA
		nullCell = self.isNull(self.ActiveCell)
		count = 0
		while count <= self.getEndColumn(self.ActiveCell)-1:
			nullCell = self.isNull(self.Offset(0 , count))
			if nullCell != self.isNull(self.ActiveCell):
				break
			count += 1
		return self.Offset(0 , count-1 if not self.isNull(self.ActiveCell) else count)

	def clcToLeft(self):				# This function is like xlToLeft in VBA
		nullCell = self.isNull(self.ActiveCell)
		count = 0
		while abs(count) <= self.getStartColumn(self.ActiveCell)-1:
			nullCell = self.isNull(self.Offset(0 , count))
			if nullCell != self.isNull(self.ActiveCell):
				break
			count -= 1
		return self.Offset(0 , count+1 if not self.isNull(self.ActiveCell) else count)

	def clcDown(self):      			# This function is like xlDown in VBA
		nullCell = self.isNull(self.ActiveCell)
		count = 0
		if nullCell:
			while nullCell:
				nullCell = self.isNull(self.Offset(count+1 , 0))
				count += 1
			return self.Offset(count , 0)
		else:
			while not nullCell:
				nullCell = self.isNull(self.Offset(count+2 , 0))
				count += 1
			return self.Offset(count , 0)

	def clcUp(self):
		pass

	def WorksheetFunctions(self, functionName , args):       # Calc.WorksheetFunctions("functionName in capital" , [args as tuple])
		smgr = XSCRIPTCONTEXT.getComponentContext().ServiceManager
		func_access = smgr.createInstanceWithContext(
			"com.sun.star.sheet.FunctionAccess",
			XSCRIPTCONTEXT.getComponentContext()
		)
		return func_access.callFunction(functionName, tuple(args))



	# Funtions for workbook class development
	def getActiveSelection(self):
		return self.active.getSelection()

	def getActiveSheet(self):
		return self.active.getActiveSheet()

	def isNull(self, oCell):
		return True if len(oCell.getString())==0 else False

	def isRangeNameGiven(self, isRowIndex=None , isColumnIndex=None , isRangeName=None):
		if isRangeName!=None and (isRowIndex==None and isColumnIndex==None):
			return True
		elif isRangeName==None and (isRowIndex!=None and isColumnIndex!=None):
			return False


	# In build funtions



Calc = WorkBook()

def Automate():
	Calc.Select(Calc.Range("A1"))
	
