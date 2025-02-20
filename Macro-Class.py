import uno

class WorkBook:
	def __init__(self):
		self.doc = XSCRIPTCONTEXT.getDocument()
		self.active = self.doc.CurrentController
		self.ActiveSheet = self.getActiveSheet()
		self.ActiveCell = self.getActiveSelection()
		
	def Select(self, oCell):
		self.active.select(oCell)
		self.ActiveCell = oCell.getCellByPosition(0 , 0)
		
	def getActiveSheet(self):
		return self.active.getActiveSheet()
		
	def getActiveSelection(self):
		return self.active.getSelection()
	
	def Cell(self, row, col):
		return self.ActiveSheet.getCellByPosition(col-1, row-1)
		
	def Range(self, rangeName=None , fromCell=None , toCell=None):
		if self.isRangeNameGiven(isRowIndex=fromCell , isColumnIndex=toCell , isRangeName=rangeName):
			return self.ActiveSheet.getCellRangeByName(cellName)
		else:
			return self.ActiveSheet.getCellRangeByName(fromCell.AbsoluteName+":"+toCell.AbsoluteName)
		
	def Offset(self, rowIndex, colIndex):
		row = self.ActiveCell.RangeAddress.StartRow + rowIndex
		col = self.ActiveCell.RangeAddress.StartColumn + colIndex
		return (self.ActiveSheet.getCellByPosition(col,row))
		
	def Row(self, oCell):
		return 	oCell.RangeAddress.StartRow + 1
	
	def Column(self, oCell):
		return oCell.RangeAddress.StartColumn + 1
	
	def MsgBox(self, message, title="Information", msgBoxType=0):
		self.active.getFrame().getComponentWindow().showMessageBox(message, title, msgBoxType)

	def isRangeNameGiven(self, isRowIndex=None , isColumnIndex=None , isRangeName=None):
		if isRangeName!=None and (isRowIndex==None and isColumnIndex==None):
			return True
		elif isRangeName==None and (isRowIndex!=None and isColumnIndex!=None):
			return False

	def isNull(self, rowIndex=None , columnIndex=None , rangeName=None):
		if self.isRangeNameGiven(isRowIndex=rowIndex , isColumnIndex=columnIndex , isRangeName=rangeName):
			return True if len(self.Range(rangeName).getString())==0 else False
		else:
			return True if len(self.Cell(rowIndex , columnIndex).getString())==0 else False




Calc = WorkBook()

def Automate():
    # edit as per your requirement
	Calc.Select(Calc.Range(fromCell=Calc.Cell(1 , 1) , toCell=Calc.Cell(5 , 5)))
	Calc.ActiveCell.setString("Hello")

