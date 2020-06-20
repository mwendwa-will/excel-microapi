import pythoncom
import numpy as np
import win32com.client

class PythonObjectLibrary:
    
    #creates a GUID to register with windows
    _reg_clsid_ = pythoncom.CreateGuid()
    
    #Registers the object as an EXE file, alternative is a dll file (IMPROC_SERVER)
    _reg_clsctx_ = pythoncom.CLSCTX_LOCAL_SERVER
    
    #This is the name of the object library
    _reg_progid_ = "Python.Object.Library"
    
    #Description of our library
    _reg_desc_ = "This is our python library"
    
    #A list of strings that indicate the public methods for the object. If they aren't listed they are considered private
    _public_methods_ =[ 'pythonSum','pythonMultiply','addArray' ]
    
    def pthonSum(self,x,y):
        return x + y
    
    def pythonMultiply(self,a,b):
        return a * b
    
    #adding a range of value
    def addArray(self,myRange):
        
        #create an instance of the range object passed through
        rngl = win32com.client.Dispatch(myRange)
        
        #convert the range into numpy array
        rnglval = np.array(list (rngl.Value))
        
        #return the sum sum of that numpy array
        return rnglval.sum()
    
if __name__ == '__main__':
        import win32com.server.register
        win32com.server.register.UseCommandLine(PythonObjectLibrary)
        
        
        
    
        
    
    
    
