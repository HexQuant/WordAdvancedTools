// dllmain.h : Declaration of module class.

class CWordAdvancedToolsModule : public ATL::CAtlDllModuleT< CWordAdvancedToolsModule >
{
public :
	DECLARE_LIBID(LIBID_WordAdvancedToolsLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_WORDADVANCEDTOOLS, "{04183b17-6eb6-4848-9c5f-1825f7d58a6a}")
};

extern class CWordAdvancedToolsModule _AtlModule;
