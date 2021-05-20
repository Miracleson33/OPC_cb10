//---------------------------------------------------------------------------

#ifndef mainH
#define mainH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <ExtCtrls.hpp>
#include <Grids.hpp>
#include <Mask.hpp>
//---------------------------------------------------------------------------
//#ifndef TrendReadThreadH
	#include <opcda.h>
	#include "opcda_i.c"
//#endif

#include <vcl.h>
#include <ExtCtrls.hpp>
#include <ComCtrls.hpp>
#include <Grids.hpp>
#include <Mask.hpp>
#include <ADODB.hpp>
#include <DB.hpp>
#include <DBClient.hpp>
#include <MConnect.hpp>
#include <SqlExpr.hpp>
#include <WideStrings.hpp>
// Include all ATL stuff:
#ifdef __BORLANDC__
#include <atlvcl.h>
#endif
#include <atlbase.h>
#include <atlcom.h>
#include <Registry.hpp>
//#include <opcenum_clsid.c>
//#include <opcenum_clsid.c>
//#include <opcenum_clsid.c>
//#include <opccomn_i.c>
//#include <opcda_cats.c>


#include <TestThread.h>
#include <TrendReadThread.h>
#include <findserv.h>
#include <trends.h>
#include <selecttrends.h>
#include <selectcolor.h>
#include "mnemo.h"
#include "startmessenger.h"
typedef interface IOPCServerList IOPCServerList;



//---------------------------------------------------------------------------
//---------------------------------------------------------------------------
//---------------------------------------------------------------------------
//---------------------------------------------------------------------------
//---------------------------------------------------------------------------
typedef void __fastcall (__closure *TOnDataChangeEvent)(
      /* [in] */ DWORD dwTransid,
      /* [in] */ OPCHANDLE hGroup,
      /* [in] */ HRESULT hrMasterquality,
      /* [in] */ HRESULT hrMastererror,
      /* [in] */ DWORD dwCount,
      /* [size_is][in] */ OPCHANDLE __RPC_FAR *phClientItems,
      /* [size_is][in] */ VARIANT __RPC_FAR *pvValues,
      /* [size_is][in] */ WORD __RPC_FAR *pwQualities,
      /* [size_is][in] */ FILETIME __RPC_FAR *pftTimeStamps,
      /* [size_is][in] */ HRESULT __RPC_FAR *pErrors);

typedef void __fastcall (__closure *TOnReadCompleteEvent)(
      /* [in] */ DWORD dwTransid,
      /* [in] */ OPCHANDLE hGroup,
      /* [in] */ HRESULT hrMasterquality,
      /* [in] */ HRESULT hrMastererror,
      /* [in] */ DWORD dwCount,
      /* [size_is][in] */ OPCHANDLE __RPC_FAR *phClientItems,
      /* [size_is][in] */ VARIANT __RPC_FAR *pvValues,
      /* [size_is][in] */ WORD __RPC_FAR *pwQualities,
      /* [size_is][in] */ FILETIME __RPC_FAR *pftTimeStamps,
	  /* [size_is][in] */ HRESULT __RPC_FAR *pErrors);


//---------------------------------------------------------------------------
class Callback : public IOPCDataCallback
{
protected:
	TOnDataChangeEvent FEvDataChange;
	TOnReadCompleteEvent FEvReadComplete;

public:
	__property TOnReadCompleteEvent EvReadComplete = {read=FEvReadComplete,
                                                     write=FEvReadComplete};
	__property TOnDataChangeEvent EvDataChange = {read=FEvDataChange,
                                                 write=FEvDataChange};
public:

	Callback()
	{
		m_ulRefs = 1;
	}

	//==========================================================================
    // IUnknown

	// QueryInterface
	STDMETHODIMP QueryInterface(REFIID iid, LPVOID* ppInterface)
	{
		if (ppInterface == NULL)
		{
			return E_INVALIDARG;
		}

		if (iid == IID_IUnknown)
		{
			*ppInterface = dynamic_cast<IUnknown*>(this);
			AddRef();
			return S_OK;
		}

		if (iid == IID_IOPCDataCallback)
		{
			*ppInterface = dynamic_cast<IOPCDataCallback*>(this);
			AddRef();
			return S_OK;
		}

		return E_NOINTERFACE;
	}

	// AddRef
	STDMETHODIMP_(ULONG) AddRef()
	{
        return InterlockedIncrement((LONG*)&m_ulRefs);
	}

	// Release
	STDMETHODIMP_(ULONG) Release()
	{
        ULONG ulRefs = InterlockedDecrement((LONG*)&m_ulRefs);

        if (ulRefs == 0)
        {
            delete this;
            return 0;
        }

        return ulRefs;
	}

	//==========================================================================
    // IOPCDataCallback

    // OnDataChange
    STDMETHODIMP OnDataChange(
        DWORD       dwTransid,
        OPCHANDLE   hGroup,
        HRESULT     hrMasterquality,
        HRESULT     hrMastererror,
        DWORD       dwCount,
        OPCHANDLE * phClientItems,
        VARIANT   * pvValues,
		WORD      * pwQualities,
        FILETIME  * pftTimeStamps,
        HRESULT   * pErrors
    )
	{
       if (FEvDataChange != NULL)
	   FEvDataChange(dwTransid, hGroup, hrMasterquality, hrMastererror,
                    dwCount, phClientItems, pvValues, pwQualities,
                    pftTimeStamps, pErrors);
/*		USES_CONVERSION;

		for (DWORD ii = 0; ii < dwCount; ii++)
		{
			VARIANT vValue;
			VariantInit(&vValue);

			if (SUCCEEDED(VariantChangeType(&vValue, &(pvValues[ii]), NULL, VT_BSTR)))
			{
//				_tprintf(_T("Handle = '%d', Value = '%s'\r\n"), phClientItems[ii], OLE2T(vValue.bstrVal));
				VariantClear(&vValue);
			}
		}
*/
		return S_OK;
	}

    // OnReadComplete
    STDMETHODIMP OnReadComplete(
        DWORD       dwTransid,
        OPCHANDLE   hGroup,
        HRESULT     hrMasterquality,
        HRESULT     hrMastererror,
        DWORD       dwCount,
        OPCHANDLE * phClientItems,
        VARIANT   * pvValues,
        WORD      * pwQualities,
        FILETIME  * pftTimeStamps,
        HRESULT   * pErrors
    )
	{
       if (FEvReadComplete != NULL)
	   FEvReadComplete(dwTransid, hGroup, hrMasterquality, hrMastererror,
                      dwCount, phClientItems, pvValues, pwQualities,
                      pftTimeStamps, pErrors);
		return S_OK;
	}

    // OnWriteComplete
    STDMETHODIMP OnWriteComplete(
        DWORD       dwTransid,
        OPCHANDLE   hGroup,
        HRESULT     hrMastererr,
        DWORD       dwCount,
        OPCHANDLE * pClienthandles,
        HRESULT   * pErrors
	)
	{
		return S_OK;
	}


    // OnCancelComplete
    STDMETHODIMP OnCancelComplete(
        DWORD       dwTransid,
        OPCHANDLE   hGroup
    )
	{
		return S_OK;
	}

private:

	ULONG m_ulRefs;
};




//---------------------------------------------------------------------------
//---------------------------------------------------------------------------
//---------------------------------------------------------------------------
//---------------------------------------------------------------------------
//---------------------------------------------------------------------------
//---------------------------------------------------------------------------
class TOPCServer
{
public:
    AnsiString NameField;
	AnsiString AboutField;
	AnsiString CLSID;
	AnsiString ComputerName;
	int idTrend;
};

class TConnectedOPCServer
{
public:
	int idTrend;
	AnsiString ComputerName;
	CComBSTR NameOPCServer;
	CComBSTR CLSIDOPCServer;
   // The OPC server
    CComPtr<IOPCServer> OPCServer;
   // BrowseServerAddressSpace
    IOPCBrowseServerAddressSpace *OPCBrowseServerAdressSpace;
   // Item flat identifier
//    CComBSTR NameOPCItem[3000];
	AnsiString NameOPCItem[3000];
	AnsiString AboutOPCItem[3000];
	AnsiString NameDBItem[3000];
	int IntValField[128];
	float FloatValField[128];
	bool BoolValField[128];
   // Item Handle
    DWORD hOPCItem[3000];
   // Handle of the OPC group
    DWORD	hOPCGroup;
   // Group parameters (see OPC doc):
    DWORD	m_dwRate;
    float	m_fDeadBand;
   // The group itself, referenced by an IUnknown interface pointer:
    CComPtr<IUnknown> OPCGroup;

	CComPtr<IOPCItemIO> OPCItemIO;
   // IOPCSyncIO interface of the group
    CComPtr<IOPCSyncIO> OPCSyncIO;
   // IOPCAsyncIO2 interface of the group
    CComPtr<IOPCAsyncIO2> OPCAsyncIO;
   // IAsyncIO2 parameters
    DWORD m_dwID; // Used as Transaction ID
    DWORD m_dwCancelID; // The cancel id for asynchronous operations

    IConnectionPoint* ipCP;
    Callback* ipCallback;
    DWORD dwAdvise;

    bool ServerConnect;
	bool ConnectPoint, BrowseServer, GroupAdd, AsyncRead, SyncRead, OnChange;
	int nItems;
};

class TListTrends
{
public:
    AnsiString NameFileConfig;
	AnsiString TableName;
	AnsiString TableNamePrefix[10];
	AnsiString TrendName;
    int NumVar;
	int MaxParam;
    bool Visible[3000];
	char Color[3000];
	int NTrend;
	int Y[3000];
	int ScaleY[3000];
	int MinY[3000];
	int MaxY[3000];
};

class TableTrends
{
public:
	AnsiString IdField;
	AnsiString DateField;
//Тип поля 0-float 1-int 2-bool
	int TypeField[128];
	AnsiString NameField[128];
	AnsiString NameFieldDB[128];
	AnsiString AboutField[128];
	int IntValField[128];
	float FloatValField[128];
	bool BoolValField[128];
};

//---------------------------------------------------------------------------
//---------------------------------------------------------------------------
//---------------------------------------------------------------------------
//---------------------------------------------------------------------------
//---------------------------------------------------------------------------
//---------------------------------------------------------------------------

class	TSBSOPCParam : public TObject
{
public:
	__fastcall TSBSOPCParam(void);
	__fastcall ~TSBSOPCParam(void);

	HANDLE	Handle;
	FILETIME	Time;
	String Name, Description, Units;
	VARIANT	Value;

// lolo, lo, hi, & hihi alarms
	float	alarms[4];
	DWORD	severity[4];
	BOOL	enabled[4];

};
///////////////////////////////
static const GUID CLSID_SBSOPCServer =
{ 0x6764a030, 0x70c, 0x11d3, { 0x80, 0xd6, 0x0, 0x60, 0x97, 0x58, 0x58, 0xbe } };

__fastcall TSBSOPCParam::TSBSOPCParam(void) : TObject()
{
	VariantInit(&Value);
	Value.vt = VT_R4;
	Value.fltVal = 0.0;
}
///////////////////////////////

__fastcall TSBSOPCParam::~TSBSOPCParam(void)
{
	VariantClear(&Value);
}

//---------------------------------------------------------------------------
class TScript
{
public:
	AnsiString CommandScript;
	AnsiString ExecuteScript;
	int NServerConnect;
	AnsiString ShowValues;
	AnsiString StoreTrends;
	AnsiString NameMnemoShow;
};
//---------------------------------------------------------------------------
class TLoadedServers
{
public:
	AnsiString NameServer;
	AnsiString CLSIDServer;
	AnsiString AboutServer;
	AnsiString ComputerNameServer;
	AnsiString NameTrendServer;
	AnsiString ConfigTrendServer;
	AnsiString PathTrendServer;
	AnsiString FileTrendServer;
	int NumVarServer;
};
//---------------------------------------------------------------------------
//-------------------------------------------------------
//---------------------------------------------------------------------------

//---------------------------------------------------------------------------
//---------------------------------------------------------------------------
//---------------------------------------------------------------------------
//---------------------------------------------------------------------------
//---------------------------------------------------------------------------
//---------------------------------------------------------------------------
class TForm1 : public TForm
{
__published:	// IDE-managed Components
		TTimer *Timer1;
	TCheckBox *CheckBox1;
	TButton *Button4;
	TGroupBox *GroupBox1;
	TLabel *Label9;
	TLabel *Label10;
	TLabel *Label12;
	TButton *Button5;
	TComboBox *ComboBox1;
	TEdit *Edit1;
	TMaskEdit *MaskEdit1;
	TCheckBox *CheckBox2;
	TGroupBox *GroupBox2;
	TLabel *Label1;
	TStringGrid *StringGrid1;
	TComboBox *ComboBox3;
	TButton *Button2;
	TButton *Button3;
	TLabel *Label3;
	TLabel *Label4;
    TLabel *Label5;
    TLabel *Label6;
    TGroupBox *GroupBox3;
    TStringGrid *StringGrid2;
    TLabel *Label2;
    TLabel *Label13;
	TButton *Button1;
	TTimer *Timer2;
	TLabel *Label14;
	TLabel *Label15;
	TButton *Button6;
	TButton *Button7;
	TCheckBox *CheckBox3;
	TButton *Button8;
	TButton *Button9;
	TLabel *Label7;
	TButton *Button10;
	TButton *Button11;
	TCheckBox *CheckBox4;
	TButton *Button12;
	TComboBox *ComboBox4;
	TLabel *Label11;
	TLabel *Label8;
    void __fastcall Button2Click(TObject *Sender);
    void __fastcall Button3Click(TObject *Sender);
    void __fastcall FormClose(TObject *Sender, TCloseAction &Action);
    void __fastcall Timer1Timer(TObject *Sender);
    void __fastcall FormCreate(TObject *Sender);
    void __fastcall Button5Click(TObject *Sender);
    void __fastcall StringGrid1SelectCell(TObject *Sender, int ACol, int ARow, bool &CanSelect);
	void __fastcall Button1Click(TObject *Sender);
    void __fastcall FormShow(TObject *Sender);
    void __fastcall ComboBox3Change(TObject *Sender);
	void __fastcall CheckBox1Click(TObject *Sender);
	void __fastcall Button4Click(TObject *Sender);
    void __fastcall StringGrid1DrawCell(TObject *Sender, int ACol,
          int ARow, TRect &Rect, TGridDrawState State);
	void __fastcall CheckBox2Click(TObject *Sender);
	void __fastcall Timer2Timer(TObject *Sender);
	void __fastcall Label14DblClick(TObject *Sender);
	void __fastcall Label15DblClick(TObject *Sender);
	void __fastcall Label12DblClick(TObject *Sender);
	void __fastcall Button6Click(TObject *Sender);
	void __fastcall ReadConfig(AnsiString ConfigFile);
	void __fastcall Button7Click(TObject *Sender);
	void __fastcall CheckBox3Click(TObject *Sender);
	void __fastcall Button8Click(TObject *Sender);
	void __fastcall Button9Click(TObject *Sender);
	void __fastcall Button10Click(TObject *Sender);
	void __fastcall Button11Click(TObject *Sender);
	void __fastcall CheckBox4Click(TObject *Sender);
	void __fastcall FormResize(TObject *Sender);
	void __fastcall Button12Click(TObject *Sender);

private:	// User declarations
public:		// User declarations
	bool ServerStarted;
	AnsiString	SBSOPCServerName, SBSOPCServerAbout, SBSHelpPath;
	TSBSOPCParam *SBSOPCParam[1000];
	int NumSBSOPCParam;

	TTrendReadThread *ThrTrendsRead;
	bool CloseTrendsReadThread;
	TDateTime LastTimeTrendsReadThread;
	TTestThread *TestThread;
	bool CloseTestThread;
	TADOQuery *TrendsDS[99];
	TADOQuery *TablesDS[99];
	TADOConnection* TrendsConn[99];
	TADOQuery *ValuesDS[99];
	TADOQuery *ParamsDS[99];
	TADOConnection* DBSQLConn[99];
//	SqlConnection* TrendsConn2[99];
	__fastcall TForm1(TComponent* Owner);
	bool AdminMode;
	TColor Colors[24];
	TColor BackColor;
	TColor BlackColor;
	TColor WhiteColor;
	TColor GrayColor;

	AnsiString NullTrendFilePath;
	AnsiString FullNullTrendFilePath;
	AnsiString TrendFilePath[99];
	AnsiString FullTrendFilePath[99];
	AnsiString TrendPath[99];
	AnsiString TrendFile[99];
	AnsiString FullTrendPath[99];
	AnsiString TrendsDirectory;
	AnsiString ConfigDirectory;
	AnsiString MnemoDirectory;
	AnsiString ProgramDirectory;

	BYTE WorkMode;
	bool ServerMode;

	bool ViewRemouteTrends;
	bool FindTrendFile[99];
	String FileTrendView;
	bool CloseTimerThread;
	bool CloseTrendsStoreThread;
	bool TrendsStoredRemoute;
	bool TrendsStoredLocal;
	int NumStoredRemoute;
	int NumStoredLocal;
	int NumListTrends;
	TListTrends ListTrends[99];
	int NumValidOPCServers;
	TOPCServer ValidOPCServers[99];
	TableTrends TTrends[99];
	int NumConnectedOPCServers;
	TConnectedOPCServer OPCServers[99];
	int SelectedOPCServer;
	int SelectedRowServer;
	int SelectedTrend;
	int SelectedParam;

	TScript Scripts[99];
	TLoadedServers LoadedServers[99];

	bool ServerConnect;
	bool ConnectPoint, BrowseServer, GroupAdd, AsyncRead, SyncRead, OnChange;
   CComPtr<IMalloc> m_ptrMalloc;
	int MaxParams;
	bool TrendsStoreRemoute;
	bool TrendsStoreLocal;
	bool ShowValues;
	TDateTime CurrentTime;
	TDateTime LastTime1s;
	TDateTime LastTime5s;
	TDateTime LastTime60s;
	bool Time1s1,Time5s1,Time60s1;
	double s1;
	double s5;
	double s60;
	double DifTime1s;
	double DifTime5s;
	double DifTime60s;
	double DifTimeTrendsReadThread;
	double DifTimeTrendsReadThread1;
	double MaxDifTimeTrendsReadThread;

	double TimeReOpen;
	double NumDayTrendFile;

	int RecordsInTable;

	CComBSTR User1;
    CComBSTR Password1;
	CComBSTR Domain1;
    CComBSTR ServerName;
   // OPC Server Prog ID
   CComBSTR NameOPCServer;
   CComBSTR CLSIDOPCServer;
   CComBSTR NameOPCServerCurrent;
   CComBSTR CLSIDOPCServerCurrent;
   // The OPC server
   CComPtr<IOPCServer> OPCServer;
   // BrowseServerAddressSpace
   IOPCBrowseServerAddressSpace *OPCBrowseServerAdressSpace;
   // Item flat identifier
   CComBSTR NameOPCItem;
   // Item Handle
   DWORD hOPCItem[3000];
   // Handle of the OPC group
   DWORD	hOPCGroup;
   // Group parameters (see OPC doc):
   DWORD	m_dwRate;
   float	m_fDeadBand;
   // The group itself, referenced by an IUnknown interface pointer:
   CComPtr<IUnknown> OPCGroup;

   CComPtr<IOPCItemIO> OPCItemIO;
   // IOPCSyncIO interface of the group
   CComPtr<IOPCSyncIO> OPCSyncIO;
   // IOPCAsyncIO2 interface of the group
   CComPtr<IOPCAsyncIO2> OPCAsyncIO;
   // IAsyncIO2 parameters
   DWORD m_dwID; // Used as Transaction ID
   DWORD m_dwCancelID; // The cancel id for asynchronous operations
   int TimeRefresh;
	bool Paused;
	IConnectionPoint* ipCP;
	Callback* ipCallback;
	DWORD dwAdvise;
	int nItems;
	int NumErrorsReads;
	int NumReads;
	int NumSaves;
	int NumBD;
	AnsiString TableVals[99];
	int idTables[99];
//	TTrendReadThread* ThrTrendsRead;
	int __fastcall ConnectToServer(int NumOPCServer);
	void __fastcall Disconnect(int NumOPCServer);
	void __fastcall ReleaseItem(int NumOPCServer);
	void __fastcall SetButtonStatus(bool Connect);
	void __fastcall GetItems(int NumOPCServer);
	void __fastcall ListItems(int NumOPCServer);
	void __fastcall ConnectToItems(int NumOPCServer);
	void __fastcall SelectOPCServer(int NumOPCServer);
	void __fastcall ShowForm(bool Advanced);
	void __fastcall ReOpenDB(int nTrendFile, bool SaveOld);
	void __fastcall ConnectDB(int nTrendFile);
	void __fastcall InitDBSQL(int nDB);
	void __fastcall ConnectDBSQL(int nDB);
	void __fastcall ReconnectDBSQL(int nDB);

	void __fastcall SaveTrendConfig();
	void __fastcall LoadTrendConfig();
	void __fastcall LoadServersConfig();
	void __fastcall ExecLoadScript();
	void __fastcall LoadServersList();

	AnsiString UserMessage;
	AnsiString LogMessage1;
	AnsiString LogMessage2;
	void __fastcall Log();

	void __fastcall OnDataChange(
	  /* [in] */ DWORD dwTransid,
      /* [in] */ OPCHANDLE hGroup,
      /* [in] */ HRESULT hrMasterquality,
	  /* [in] */ HRESULT hrMastererror,
	  /* [in] */ DWORD dwCount,
      /* [size_is][in] */ OPCHANDLE __RPC_FAR *phClientItems,
	  /* [size_is][in] */ VARIANT __RPC_FAR *pvValues,
	  /* [size_is][in] */ WORD __RPC_FAR *pwQualities,
	  /* [size_is][in] */ FILETIME __RPC_FAR *pftTimeStamps,
	  /* [size_is][in] */ HRESULT __RPC_FAR *pErrors);

	void __fastcall OnReadComplete(
	  /* [in] */ DWORD dwTransid,
      /* [in] */ OPCHANDLE hGroup,
      /* [in] */ HRESULT hrMasterquality,
	  /* [in] */ HRESULT hrMastererror,
	  /* [in] */ DWORD dwCount,
      /* [size_is][in] */ OPCHANDLE __RPC_FAR *phClientItems,
	  /* [size_is][in] */ VARIANT __RPC_FAR *pvValues,
	  /* [size_is][in] */ WORD __RPC_FAR *pwQualities,
      /* [size_is][in] */ FILETIME __RPC_FAR *pftTimeStamps,
	  /* [size_is][in] */ HRESULT __RPC_FAR *pErrors);
};



//---------------------------------------------------------------------------
extern PACKAGE TForm1 *Form1;
//---------------------------------------------------------------------------
EXTERN_C const IID IID_IOPCServerList;
typedef GUID CATID;
#if defined(__cplusplus) && !defined(CINTERFACE)

    interface DECLSPEC_UUID("13486D50-4821-11D2-A494-3CB306C10000")
	IOPCServerList : public IUnknown
    {
	public:
        virtual HRESULT STDMETHODCALLTYPE EnumClassesOfCategories(
            /* [in] */ ULONG cImplemented,
			/* [size_is][in] */ CATID __RPC_FAR rgcatidImpl[  ],
            /* [in] */ ULONG cRequired,
            /* [size_is][in] */ CATID __RPC_FAR rgcatidReq[  ],
			/* [out] */ IEnumGUID __RPC_FAR *__RPC_FAR *ppenumClsid) = 0;

        virtual HRESULT STDMETHODCALLTYPE GetClassDetails(
			/* [in] */ REFCLSID clsid,
            /* [out] */ LPOLESTR __RPC_FAR *ppszProgID,
            /* [out] */ LPOLESTR __RPC_FAR *ppszUserType) = 0;

        virtual HRESULT STDMETHODCALLTYPE CLSIDFromProgID(
            /* [in] */ LPCOLESTR szProgId,
			/* [out] */ LPCLSID clsid) = 0;

    };

 /*   typedef struct _IID
    {
		unsigned long x;
        unsigned short s1;
        unsigned short s2;
		unsigned char  c[8];
    } IID;     */

 #define MIDL_DEFINE_GUID(type,name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8) \
        const type name = {l,w1,w2,{b1,b2,b3,b4,b5,b6,b7,b8}}

    //const IID IID_IOPCServerList = {0x13486D50,0x4821,0x11D2,{0xA4,0x94,0x3C,0xB3,0x06,0xC1,0x00,0x00}};
    MIDL_DEFINE_GUID(IID, IID_IOPCServerList,0x13486D50,0x4821,0x11D2,0xA4,0x94,0x3C,0xB3,0x06,0xC1,0x00,0x00);
    MIDL_DEFINE_GUID(IID, IID_IOPCServerList2,0x9DD0B56C,0xAD9E,0x43ee,0x83,0x05,0x48,0x7F,0x31,0x88,0xBF,0x7A);

    //const CLSID CLSID_OPCServerList = {0x13486D51,0x4821,0x11D2,{0xA4,0x94,0x3C,0xB3,0x06,0xC1,0x00,0x00}};
    MIDL_DEFINE_GUID(CLSID, CLSID_OpcServerList,0x13486D51,0x4821,0x11D2,0xA4,0x94,0x3C,0xB3,0x06,0xC1,0x00,0x00);

/*    const GUID CATID_OPCDAServer10 =
    { 0x63d5f430, 0xcfe4, 0x11d1, { 0xb2, 0xc8, 0x0, 0x60, 0x8, 0x3b, 0xa1, 0xfb } };
// {63D5F430-CFE4-11d1-B2C8-0060083BA1FB}

    const GUID CATID_OPCDAServer20 =
    { 0x63d5f432, 0xcfe4, 0x11d1, { 0xb2, 0xc8, 0x0, 0x60, 0x8, 0x3b, 0xa1, 0xfb } };
// {63D5F432-CFE4-11d1-B2C8-0060083BA1FB}      */

    MIDL_DEFINE_GUID(IID, IID_CATID_OPCDAServer10,0x63D5F430,0xCFE4,0x11d1,0xB2,0xC8,0x00,0x60,0x08,0x3B,0xA1,0xFB);
    MIDL_DEFINE_GUID(IID, IID_CATID_OPCDAServer20,0x63D5F432,0xCFE4,0x11d1,0xB2,0xC8,0x00,0x60,0x08,0x3B,0xA1,0xFB);
    MIDL_DEFINE_GUID(IID, IID_CATID_OPCDAServer30,0xCC603642,0x66D7,0x48f1,0xB6,0x9A,0xB6,0x25,0xE7,0x36,0x52,0xD7);

#endif
#endif
