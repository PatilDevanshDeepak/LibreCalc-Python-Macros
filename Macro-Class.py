import uno

class WorkBook:
	def __init__(self):
		self.doc = XSCRIPTCONTEXT.getDocument()
		self.active = self.doc.CurrentController
		self.ActiveSheet = self.getActiveSheet()
		self.ActiveCell = self.getActiveCell()
		
	def Select(self, oCell):
		self.active.select(oCell)
		self.ActiveCell = self.getActiveCell() 
		
	def getActiveSheet(self):
		return self.active.getActiveSheet()
		
	def getActiveCell(self):
		return self.active.getSelection()
	
	def Cell(self, row, col):
		return (self.ActiveSheet.getCellByPosition(col-1, row-1))
		
	def Range(self, rangeName):
		return (self.ActiveSheet.getCellRangeByName(rangeName))
		
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


Calc = WorkBook()

def Automate():
    # edit as per your requirement
	Calc.Select(Calc.Cell(1 , 1)) # Select Cell A1
	Calc.ActiveCell.setString("Hello World")

