import comtypes.client
import comtypes.server.localserver
# comtypes will generate TypeLib wrappers that are cached on disk. These are unfortunately not
# regenerated if a COM interface is changed, which can lead to problems. Therefore, please eiher:
# * Delete the "comtypes.client.gen_dir" folder when changing a COM interface without updating the GUID, or
# * Set "comtypes.client.gen_dir = None" during script startup to disable caching (doesn't always work).

from packaging.version import parse
if parse(comtypes.__version__) < parse("1.4.9"):
    raise ImportError("comtypes 1.4.9 or newer required due to https://github.com/enthought/comtypes/pull/678")

# Load type library
MyInterfaces = comtypes.client.GetModule("../x64/Debug/MyInterfaces.tlb")
#MyInterfaces = comtypes.client.GetModule("../x64/Release/MyInterfaces.tlb")

class NumberCruncher(comtypes.CoClass):
    '''Non-creatable COM class that doesn't need any CLSID'''
    _com_interfaces_ = [MyInterfaces.INumberCruncher]

    def ComputePi(self):
        return 3.14
    

class MyServerImpl(MyInterfaces.MyServer):
     # registry entries
    _reg_threading_ = "Both"
    _reg_desc_ = "Python-based COM server for testing"
    _reg_clsctx_ = comtypes.CLSCTX_LOCAL_SERVER # run in separate process
    _regcls_ = comtypes.server.localserver.REGCLS_MULTIPLEUSE
    
    def GetNumberCruncher(self):
        obj = NumberCruncher()
        return obj.QueryInterface(MyInterfaces.INumberCruncher)
    
    def Subscribe(self, client):
        pass # not implemented
    
    def Unsubscribe(self, client):
        pass # not implemented


if __name__=="__main__":
    from comtypes.server.register import UseCommandLine
    UseCommandLine(MyServerImpl) # will parse /regserver and /unregserver arguments
