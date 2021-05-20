//---------------------------------------------------------------------------

#include <vcl.h>
#include <assert.h>

#pragma hdrstop

#include "main.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TForm1 *Form1;
#define MAX_KEYLEN 256
extern const GUID CATID_OPCDAServer10;
extern const GUID CATID_OPCDAServer20;
extern const GUID CATID_OPCDAServer30;
extern const CLSID CLSID_StdComponentCategoriesMgr;
//extern const CLSID CLSID_OPCServerList;
extern const IID IID_ICatInformation;
extern const IID IID_IOPCServerList;

// Ole initializer/deinitializer
static struct SOLEINIT
{
   SOLEINIT()
   {
	  CoInitialize(NULL);
   }
   ~SOLEINIT()
   {
	  CoUninitialize();
   }
} oleinit__;
///////////////////////////////

///////////////////////////////
///////////////////////////////Сервер(
///////////////////////////////
//#include "Unit1.h"
//#define STRICT 1
//#include "WTOPCsvrAPI.h"
//#pragma comment(lib, "wtopcsvr1.lib")
//#include "Unit2.h"
//#include "Unit3.h"

/*
class	CTag : public TObject
{
public:
	__fastcall CTag(void);
	__fastcall ~CTag(void);

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

__fastcall CTag::CTag(void) : TObject()
{
	VariantInit(&Value);
	Value.vt = VT_R4;
	Value.fltVal = 0.0;
}
///////////////////////////////

__fastcall CTag::~CTag(void)
{
	VariantClear(&Value);
}
		 */
void CALLBACK _export UnknownTagProc(LPSTR Path, LPSTR Name);
void CALLBACK _export TagRemovedProc(HANDLE hTag, LPSTR Path, LPSTR Name);
void CALLBACK _export WriteNotifyProc(HANDLE Handle, VARIANT *pNewValue, DWORD *pDevError);
void CALLBACK _export DeviceReadProc(HANDLE Handle, VARIANT *pNewValue, WORD *pQuality, FILETIME *pTimestamp);
void CALLBACK _export DisconnectProc(DWORD NumbrActiveClients);

///////////////////////////////
///////////////////////////////)Сервер
///////////////////////////////

inline AnsiString Variant2Str(VARIANT& v)
{
    Variant var(v);
    return VarToStr(var);
}

//---------------------------------------------------------------------------
__fastcall TForm1::TForm1(TComponent* Owner)
		: TForm(Owner)
{
//	startform->ShowModal();
	ProgramDirectory=ExtractFilePath(Application->ExeName);
	ServerStarted=false;
	SBSHelpPath = ExtractFilePath(Application->ExeName);
	SBSOPCServerName = "SBSOPCServer";
	SBSOPCServerAbout = "SBS OPC Server to data transmission from SBS OPC Client";

	NumBD=7;
	DWORD l=16;
	LPTSTR User="1234567890123456";
	GetUserName(User, &l);
	UserMessage.operator +=(User);
	LogMessage1=L"-+-+-+-+-+-+-+-+-+-";
	LogMessage2=L"-+-+-+-+-+-+-+-+-+-";
	Log();
	LogMessage1=L"Запуск программы";
	LogMessage2=L"начало инициализации";
	Log();

	UnicodeString ustr;
	UnicodeString ustr1;
	int lustr;
	int lustr1;

	ServerMode=false;
	WorkMode=0;

	ustr=Application->ExeName;
	ustr1="opccbks.exe";
	lustr=ustr.Length();
	lustr1=ustr1.Length();
	ustr=ustr.SubString(lustr-lustr1+1,lustr1);
	if (!ustr.CompareIC(ustr1))
	{
		ServerMode=true;
		ReadConfig(".\\opccbks.ini");
	}

	ustr=Application->ExeName;
	ustr1="opccbkc.exe";
	lustr=ustr.Length();
	lustr1=ustr1.Length();
	ustr=ustr.SubString(lustr-lustr1+1,lustr1);
	if (!ustr.CompareIC(ustr1))
	{
		ServerMode=false;
		ReadConfig(".\\opccbkc.ini");
	}

	ustr=Application->ExeName;
	ustr1="opccbts.exe";
	lustr=ustr.Length();
	lustr1=ustr1.Length();
	ustr=ustr.SubString(lustr-lustr1+1,lustr1);
	if (!ustr.CompareIC(ustr1))
	{
		ServerMode=true;
		ReadConfig(".\\opccbts.ini");
	}

	ustr=Application->ExeName;
	ustr1="opccbtc.exe";
	lustr=ustr.Length();
	lustr1=ustr1.Length();
	ustr=ustr.SubString(lustr-lustr1+1,lustr1);
	if (!ustr.CompareIC(ustr1))
	{
		ServerMode=false;
		ReadConfig(".\\opccbtc.ini");
	}

	ConfigDirectory=ExtractFilePath(Application->ExeName)+"config\\";
	CreateDirectory(ConfigDirectory.c_str(), NULL);
	TrendsDirectory=ExtractFilePath(Application->ExeName)+"trends\\";
	CreateDirectory(TrendsDirectory.c_str(), NULL);
	MnemoDirectory=ExtractFilePath(Application->ExeName)+"mnemo\\";
	CreateDirectory(MnemoDirectory.c_str(), NULL);

	NullTrendFilePath="base.000";
	FullNullTrendFilePath=ExtractFilePath(Application->ExeName);
	FullNullTrendFilePath.operator +=(NullTrendFilePath);

	LoadServersList();

	WIN32_FIND_DATA FindFileData;
	HANDLE hFind;

	for (int i=0; i < 10; i++)
	{
		FullTrendPath[i]=ExtractFilePath(Application->ExeName);
		FullTrendPath[i].operator +=(TrendPath[i]);
		FullTrendFilePath[i]=ExtractFilePath(Application->ExeName);
		FullTrendFilePath[i].operator +=(TrendFilePath[i]);
		hFind = FindFirstFile(TrendFilePath[i].c_str(), &FindFileData);
		if (hFind == INVALID_HANDLE_VALUE)
		{
			FindTrendFile[i]=false;
		}
		else
		{
			FindTrendFile[i]=true;
			FindClose(hFind);
		}
	}

	AdminMode=false;
	ComboBox1->Text="ss1-1";
//	ComboBox4->Text="OPCSERVERASUTP";
	ComboBox4->Text="SS2-4";
//	ComboBox4->Text="TRENDSERVER";
//	Edit1->Text="asutp";
//	MaskEdit1->Text="warwarwar";
	Edit1->Text="userasutp";
	MaskEdit1->Text="1qazxsw2";


	RecordsInTable=131072;//3456;//1728;
	TrendsStoreRemoute=false;
	TrendsStoreLocal=false;
	ShowValues=true;
	s1=(double)((double)(1)/(double)(24*60*60));
	s5=(double)((double)(5)/(double)(24*60*60));
	s60=(double)((double)(60)/(double)(24*60*60));
	DifTime1s=0;
	DifTime5s=0;
	DifTime60s=0;

	StringGrid1->ColWidths[0]=0;
	StringGrid1->ColWidths[1]=StringGrid1->Width-85;
	StringGrid1->ColWidths[2]=79;

	NumErrorsReads=0;
	NumReads=0;
	NumSaves=0;
//	StringGrid2->ColWidths[0]=StringGrid2->Width/2;
//	StringGrid2->ColWidths[1]=StringGrid2->Width-StringGrid2->ColWidths[0]-5;
	StringGrid2->ColCount=3;
	StringGrid2->ColWidths[0]=StringGrid2->Width*5;//0;//StringGrid2->Width/2;
	StringGrid2->ColWidths[0]/=12;//0;//StringGrid2->Width/2;
	StringGrid2->ColWidths[1]=StringGrid2->Width/4;
	StringGrid2->ColWidths[2]=StringGrid2->Width-StringGrid2->ColWidths[0]-StringGrid2->ColWidths[1]-24;

	CheckBox2->Checked=true;
	ShowForm(CheckBox2->Checked);

	CoGetMalloc(MEMCTX_TASK, &m_ptrMalloc);

//	StringGrid1->RowCount=NumConnectedOPCServers;
	int n=0;
	for (int i=0;i<NumConnectedOPCServers;i++)
	{
		n++;
		StringGrid1->RowCount=n;
		StringGrid1->Cells[0][n-1]=AnsiString::AnsiString(OPCServers[i].CLSIDOPCServer);
		StringGrid1->Cells[1][n-1]=AnsiString::AnsiString(OPCServers[i].NameOPCServer);
		StringGrid1->Cells[2][n-1]=AnsiString::AnsiString(OPCServers[i].ComputerName);
		ComboBox3->Items->Add(AnsiString::AnsiString(OPCServers[i].CLSIDOPCServer));
	}
	TTrends[0].IdField="id";
	TTrends[0].DateField="date";

	TTrends[0].TypeField[0]=0;
	TTrends[0].NameField[0]="FгазКотел";
	TTrends[0].AboutField[0]="FгазКотел";
	TTrends[0].TypeField[1]=0;
	TTrends[0].NameField[1]="FнПродувки";
	TTrends[0].AboutField[1]="FнПродувки";
	TTrends[0].TypeField[2]=0;
	TTrends[0].NameField[2]="FпПар";
	TTrends[0].AboutField[2]="FпПар";
	TTrends[0].TypeField[3]=0;
	TTrends[0].NameField[3]="FпитВоды";
	TTrends[0].AboutField[3]="FпитВоды";
	TTrends[0].TypeField[4]=0;
	TTrends[0].NameField[4]="SсреднееДф";
	TTrends[0].AboutField[4]="SсреднееДф";
	TTrends[0].TypeField[5]=2;
	TTrends[0].NameField[5]="ЕстьФакелК5";
	TTrends[0].AboutField[5]="ЕстьФакелК5";
	TTrends[0].TypeField[6]=0;
	TTrends[0].NameField[6]="Н_Барабан";
	TTrends[0].AboutField[6]="Н_Барабан";
	TTrends[0].TypeField[7]=0;
	TTrends[0].NameField[7]="О2среднее";
	TTrends[0].AboutField[7]="О2среднее";
	TTrends[0].TypeField[8]=0;
	TTrends[0].NameField[8]="О2топкаЛ1";
	TTrends[0].AboutField[8]="О2топкаЛ1";
	TTrends[0].TypeField[9]=0;
	TTrends[0].NameField[9]="О2топкаП1";
	TTrends[0].AboutField[9]="О2топкаП1";
	TTrends[0].TypeField[10]=0;
	TTrends[0].NameField[10]="ПолГРК";
	TTrends[0].AboutField[10]="ПолГРК";
	TTrends[0].TypeField[11]=0;
	TTrends[0].NameField[11]="ПолПК255";
	TTrends[0].AboutField[11]="ПолПК255";
	TTrends[0].TypeField[12]=0;
	TTrends[0].NameField[12]="ПолРК255";
	TTrends[0].AboutField[12]="ПолРК255";
	TTrends[0].TypeField[13]=0;
	TTrends[0].NameField[13]="РвБарабане";
	TTrends[0].AboutField[13]="РвБарабане";
	TTrends[0].TypeField[14]=0;
	TTrends[0].NameField[14]="Рвозд_за_ВЗП_Л";
	TTrends[0].AboutField[14]="Рвозд_за_ВЗП_Л";
	TTrends[0].TypeField[15]=0;
	TTrends[0].NameField[15]="Рвозд_за_ВЗП_П";
	TTrends[0].AboutField[15]="Рвозд_за_ВЗП_П";
	TTrends[0].TypeField[16]=0;
	TTrends[0].NameField[16]="РвоздСреднДф";
	TTrends[0].AboutField[16]="РвоздСреднДф";
	TTrends[0].TypeField[17]=0;
	TTrends[0].NameField[17]="Ргаз_за_РК";
	TTrends[0].AboutField[17]="Ргаз_за_РК";
	TTrends[0].TypeField[18]=0;
	TTrends[0].NameField[18]="РгазДоОК";
	TTrends[0].AboutField[18]="РгазДоОК";
	TTrends[0].TypeField[19]=0;
	TTrends[0].NameField[19]="РпПар";
	TTrends[0].AboutField[19]="РпПар";
	TTrends[0].TypeField[20]=0;
	TTrends[0].NameField[20]="РпвДоСУП";
	TTrends[0].AboutField[20]="РпвДоСУП";
	TTrends[0].TypeField[21]=0;
	TTrends[0].NameField[21]="ТвздЗаВЗП2Л1";
	TTrends[0].AboutField[21]="ТвздЗаВЗП2Л1";
	TTrends[0].TypeField[22]=0;
	TTrends[0].NameField[22]="ТвздЗаВЗП2П1";
	TTrends[0].AboutField[22]="ТвздЗаВЗП2П1";
	TTrends[0].TypeField[23]=0;
	TTrends[0].NameField[23]="ТвздНарЛ1";
	TTrends[0].AboutField[23]="ТвздНарЛ1";
	TTrends[0].TypeField[24]=0;
	TTrends[0].NameField[24]="ТвздНарП1";
	TTrends[0].AboutField[24]="ТвздНарП1";
	TTrends[0].TypeField[25]=0;
	TTrends[0].NameField[25]="ТгазНаКотел";
	TTrends[0].AboutField[25]="ТгазНаКотел";
	TTrends[0].TypeField[26]=0;
	TTrends[0].NameField[26]="ТзаКалорЛ1";
	TTrends[0].AboutField[26]="ТзаКалорЛ1";
	TTrends[0].TypeField[27]=0;
	TTrends[0].NameField[27]="ТзаКалорП1";
	TTrends[0].AboutField[27]="ТзаКалорП1";
	TTrends[0].TypeField[28]=0;
	TTrends[0].NameField[28]="ТпарЗаКотл";
	TTrends[0].AboutField[28]="ТпарЗаКотл";
	TTrends[0].TypeField[29]=0;
	TTrends[0].NameField[29]="ТпитВоды";
	TTrends[0].AboutField[29]="ТпитВоды";
	TTrends[0].TypeField[30]=0;
	TTrends[0].NameField[30]="ТуходГазовЛ1";
	TTrends[0].AboutField[30]="ТуходГазовЛ1";
	TTrends[0].TypeField[31]=0;
	TTrends[0].NameField[31]="ТуходГазовП1";
	TTrends[0].AboutField[31]="ТуходГазовП1";
	TTrends[0].TypeField[32]=0;
	TTrends[0].NameField[32]="ТшкУСОк5";
	TTrends[0].AboutField[32]="ТшкУСОк5";
	TTrends[0].TypeField[33]=1;
	TTrends[0].NameField[33]="ЧасыРаботыК5";
	TTrends[0].AboutField[33]="ЧасыРаботыК5";
	TTrends[0].TypeField[34]=1;
	TTrends[0].NameField[34]="ЧасыРаботыК5М";
	TTrends[0].AboutField[34]="ЧасыРаботыК5М";
	TTrends[0].TypeField[35]=0;
	TTrends[0].NameField[35]="МаксВремяОпроса";
	TTrends[0].AboutField[35]="МаксВремяОпроса";

	TTrends[1].IdField="id";
	TTrends[1].DateField="date";

	TTrends[1].TypeField[0]=0;
	TTrends[1].NameField[0]="FгазКотел";
	TTrends[1].AboutField[0]="FгазКотел";
	TTrends[1].TypeField[1]=0;
	TTrends[1].NameField[1]="FнПродувки";
	TTrends[1].AboutField[1]="FнПродувки";
	TTrends[1].TypeField[2]=0;
	TTrends[1].NameField[2]="FпПар";
	TTrends[1].AboutField[2]="FпПар";
	TTrends[1].TypeField[3]=0;
	TTrends[1].NameField[3]="FпитВоды";
	TTrends[1].AboutField[3]="FпитВоды";
	TTrends[1].TypeField[4]=0;
	TTrends[1].NameField[4]="SвТопкеСредн";
	TTrends[1].AboutField[4]="SвТопкеСредн";
	TTrends[1].TypeField[5]=2;
	TTrends[1].NameField[5]="ЕстьФакел";
	TTrends[1].AboutField[5]="ЕстьФакел";
	TTrends[1].TypeField[6]=0;
	TTrends[1].NameField[6]="Н_Барабан";
	TTrends[1].AboutField[6]="Н_Барабан";
	TTrends[1].TypeField[7]=0;
	TTrends[1].NameField[7]="О2среднее";
	TTrends[1].AboutField[7]="О2среднее";
	TTrends[1].TypeField[8]=0;
	TTrends[1].NameField[8]="О2топкаЛ1";
	TTrends[1].AboutField[8]="О2топкаЛ1";
	TTrends[1].TypeField[9]=0;
	TTrends[1].NameField[9]="О2топкаП1";
	TTrends[1].AboutField[9]="О2топкаП1";
	TTrends[1].TypeField[10]=0;
	TTrends[1].NameField[10]="ПолГРК";
	TTrends[1].AboutField[10]="ПолГРК";
	TTrends[1].TypeField[11]=0;
	TTrends[1].NameField[11]="ПолПК255";
	TTrends[1].AboutField[11]="ПолПК255";
	TTrends[1].TypeField[12]=0;
	TTrends[1].NameField[12]="ПолРК255";
	TTrends[1].AboutField[12]="ПолРК255";
	TTrends[1].TypeField[13]=0;
	TTrends[1].NameField[13]="РвБарабане";
	TTrends[1].AboutField[13]="РвБарабане";
	TTrends[1].TypeField[14]=0;
	TTrends[1].NameField[14]="РвоздЗаВЗП_Л";
	TTrends[1].AboutField[14]="РвоздЗаВЗП_Л";
	TTrends[1].TypeField[15]=0;
	TTrends[1].NameField[15]="РвоздЗаВЗП_П";
	TTrends[1].AboutField[15]="РвоздЗаВЗП_П";
	TTrends[1].TypeField[16]=0;
	TTrends[1].NameField[16]="РвоздСредн_Д";
	TTrends[1].AboutField[16]="РвоздСредн_Д";
	TTrends[1].TypeField[17]=0;
	TTrends[1].NameField[17]="РгазДоОК";
	TTrends[1].AboutField[17]="РгазДоОК";
	TTrends[1].TypeField[18]=0;
	TTrends[1].NameField[18]="РгазЗаГРК";
	TTrends[1].AboutField[18]="РгазЗаГРК";
	TTrends[1].TypeField[19]=0;
	TTrends[1].NameField[19]="РпПарК6";
	TTrends[1].AboutField[19]="РпПарК6";
	TTrends[1].TypeField[20]=0;
	TTrends[1].NameField[20]="РпвДоСУП";
	TTrends[1].AboutField[20]="РпвДоСУП";
	TTrends[1].TypeField[21]=0;
	TTrends[1].NameField[21]="Т_УСО01";
	TTrends[1].AboutField[21]="Т_УСО01";
	TTrends[1].TypeField[22]=0;
	TTrends[1].NameField[22]="ТвздЗаВЗП2Л1";
	TTrends[1].AboutField[22]="ТвздЗаВЗП2Л1";
	TTrends[1].TypeField[23]=0;
	TTrends[1].NameField[23]="ТвздЗаВЗП2П1";
	TTrends[1].AboutField[23]="ТвздЗаВЗП2П1";
	TTrends[1].TypeField[24]=0;
	TTrends[1].NameField[24]="ТвздНарЛ1";
	TTrends[1].AboutField[24]="ТвздНарЛ1";
	TTrends[1].TypeField[25]=0;
	TTrends[1].NameField[25]="ТвздНарП1";
	TTrends[1].AboutField[25]="ТвздНарП1";
	TTrends[1].TypeField[26]=0;
	TTrends[1].NameField[26]="ТгазНаКотел";
	TTrends[1].AboutField[26]="ТгазНаКотел";
	TTrends[1].TypeField[27]=0;
	TTrends[1].NameField[27]="ТзаКалорЛ1";
	TTrends[1].AboutField[27]="ТзаКалорЛ1";
	TTrends[1].TypeField[28]=0;
	TTrends[1].NameField[28]="ТзаКалорП1";
	TTrends[1].AboutField[28]="ТзаКалорП1";
	TTrends[1].TypeField[29]=0;
	TTrends[1].NameField[29]="ТпарЗаКотл6";
	TTrends[1].AboutField[29]="ТпарЗаКотл6";
	TTrends[1].TypeField[30]=0;
	TTrends[1].NameField[30]="ТпитВоды";
	TTrends[1].AboutField[30]="ТпитВоды";
	TTrends[1].TypeField[31]=0;
	TTrends[1].NameField[31]="ТуходГазовЛ1";
	TTrends[1].AboutField[31]="ТуходГазовЛ1";
	TTrends[1].TypeField[32]=0;
	TTrends[1].NameField[32]="ТуходГазовП1";
	TTrends[1].AboutField[32]="ТуходГазовП1";
	TTrends[1].TypeField[33]=1;
	TTrends[1].NameField[33]="ЧасыРаботыК6";
	TTrends[1].AboutField[33]="ЧасыРаботыК6";
	TTrends[1].TypeField[34]=1;
	TTrends[1].NameField[34]="ЧасыРаботыК6М";
	TTrends[1].AboutField[34]="ЧасыРаботыК6М";

	TTrends[2].IdField="id";
	TTrends[2].DateField="date";

	TTrends[2].TypeField[0]=0;
	TTrends[2].NameField[0]="FгазКотел";
	TTrends[2].AboutField[0]="FгазКотел";
	TTrends[2].TypeField[1]=0;
	TTrends[2].NameField[1]="FпПар";
	TTrends[2].AboutField[1]="FпПар";
	TTrends[2].TypeField[2]=0;
	TTrends[2].NameField[2]="SвТопкеСредн";
	TTrends[2].AboutField[2]="SвТопкеСредн";
	TTrends[2].TypeField[3]=2;
	TTrends[2].NameField[3]="ЕстьФакел";
	TTrends[2].AboutField[3]="ЕстьФакел";
	TTrends[2].TypeField[4]=0;
	TTrends[2].NameField[4]="Н_Барабан";
	TTrends[2].AboutField[4]="Н_Барабан";
	TTrends[2].TypeField[5]=0;
	TTrends[2].NameField[5]="О2среднее";
	TTrends[2].AboutField[5]="О2среднее";
	TTrends[2].TypeField[6]=0;
	TTrends[2].NameField[6]="О2топкаЛ1";
	TTrends[2].AboutField[6]="О2топкаЛ1";
	TTrends[2].TypeField[7]=0;
	TTrends[2].NameField[7]="О2топкаП1";
	TTrends[2].AboutField[7]="О2топкаП1";
	TTrends[2].TypeField[8]=0;
	TTrends[2].NameField[8]="ПолГРК";
	TTrends[2].AboutField[8]="ПолГРК";
	TTrends[2].TypeField[9]=0;
	TTrends[2].NameField[9]="РвоздЗаВЗПлев";
	TTrends[2].AboutField[9]="РвоздЗаВЗПлев";
	TTrends[2].TypeField[10]=0;
	TTrends[2].NameField[10]="РвоздЗаВЗПпр";
	TTrends[2].AboutField[10]="РвоздЗаВЗПпр";
	TTrends[2].TypeField[11]=0;
	TTrends[2].NameField[11]="РвоздСредн_Д";
	TTrends[2].AboutField[11]="РвоздСредн_Д";
	TTrends[2].TypeField[12]=0;
	TTrends[2].NameField[12]="РгазДоОК";
	TTrends[2].AboutField[12]="РгазДоОК";
	TTrends[2].TypeField[13]=0;
	TTrends[2].NameField[13]="РгазЗаРК";
	TTrends[2].AboutField[13]="РгазЗаРК";
	TTrends[2].TypeField[14]=0;
	TTrends[2].NameField[14]="РпПарК7";
	TTrends[2].AboutField[14]="РпПарК7";
	TTrends[2].TypeField[15]=0;
	TTrends[2].NameField[15]="ТгазНаКотел";
	TTrends[2].AboutField[15]="ТгазНаКотел";
	TTrends[2].TypeField[16]=0;
	TTrends[2].NameField[16]="ТпарЗаКотл7";
	TTrends[2].AboutField[16]="ТпарЗаКотл7";
	TTrends[2].TypeField[17]=1;
	TTrends[2].NameField[17]="ЧасыРаботыК7";
	TTrends[2].AboutField[17]="ЧасыРаботыК7";
	TTrends[2].TypeField[18]=1;
	TTrends[2].NameField[18]="ЧасыРаботыК7М";
	TTrends[2].AboutField[18]="ЧасыРаботыК7М";
	TTrends[2].TypeField[19]=0;
	TTrends[2].NameField[19]="FнПродувкиСум";
	TTrends[2].AboutField[19]="FнПродувкиСум";
	TTrends[2].TypeField[20]=0;
	TTrends[2].NameField[20]="FпитВоды";
	TTrends[2].AboutField[20]="FпитВоды";
	TTrends[2].TypeField[21]=0;
	TTrends[2].NameField[21]="ПолРК255";
	TTrends[2].AboutField[21]="ПолРК255";
	TTrends[2].TypeField[22]=0;
	TTrends[2].NameField[22]="РвБарабане";
	TTrends[2].AboutField[22]="РвБарабане";
	TTrends[2].TypeField[23]=0;
	TTrends[2].NameField[23]="РпвДоСУП";
	TTrends[2].AboutField[23]="РпвДоСУП";
	TTrends[2].TypeField[24]=0;
	TTrends[2].NameField[24]="ТвздЗаВЗП2Л1";
	TTrends[2].AboutField[24]="ТвздЗаВЗП2Л1";
	TTrends[2].TypeField[25]=0;
	TTrends[2].NameField[25]="ТвздЗаВЗП2П1";
	TTrends[2].AboutField[25]="ТвздЗаВЗП2П1";
	TTrends[2].TypeField[26]=0;
	TTrends[2].NameField[26]="ТвздНарЛ1";
	TTrends[2].AboutField[26]="ТвздНарЛ1";
	TTrends[2].TypeField[27]=0;
	TTrends[2].NameField[27]="ТвздНарП1";
	TTrends[2].AboutField[27]="ТвздНарП1";
	TTrends[2].TypeField[28]=0;
	TTrends[2].NameField[28]="ТзаКалорЛ1";
	TTrends[2].AboutField[28]="ТзаКалорЛ1";
	TTrends[2].TypeField[29]=0;
	TTrends[2].NameField[29]="ТзаКалорП1";
	TTrends[2].AboutField[29]="ТзаКалорП1";
	TTrends[2].TypeField[30]=0;
	TTrends[2].NameField[30]="ТпитВоды";
	TTrends[2].AboutField[30]="ТпитВоды";
	TTrends[2].TypeField[31]=0;
	TTrends[2].NameField[31]="ТуходГазовЛ1";
	TTrends[2].AboutField[31]="ТуходГазовЛ1";
	TTrends[2].TypeField[32]=0;
	TTrends[2].NameField[32]="ТуходГазовП1";
	TTrends[2].AboutField[32]="ТуходГазовП1";
	TTrends[2].TypeField[33]=0;
	TTrends[2].NameField[33]="ТшкУСО";
	TTrends[2].AboutField[33]="ТшкУСО";

	TTrends[3].IdField="id";
	TTrends[3].DateField="date";

	TTrends[3].TypeField[0]=0;
	TTrends[3].NameField[0]="FгазКотел";
	TTrends[3].AboutField[0]="FгазКотел";
	TTrends[3].TypeField[1]=0;
	TTrends[3].NameField[1]="FпПар";
	TTrends[3].AboutField[1]="FпПар";
	TTrends[3].TypeField[2]=0;
	TTrends[3].NameField[2]="SвТопкеСредн";
	TTrends[3].AboutField[2]="SвТопкеСредн";
	TTrends[3].TypeField[3]=2;
	TTrends[3].NameField[3]="ЕстьФакел";
	TTrends[3].AboutField[3]="ЕстьФакел";
	TTrends[3].TypeField[4]=0;
	TTrends[3].NameField[4]="Н_Барабан";
	TTrends[3].AboutField[4]="Н_Барабан";
	TTrends[3].TypeField[5]=0;
	TTrends[3].NameField[5]="О2среднее";
	TTrends[3].AboutField[5]="О2среднее";
	TTrends[3].TypeField[6]=0;
	TTrends[3].NameField[6]="О2топкаЛ1";
	TTrends[3].AboutField[6]="О2топкаЛ1";
	TTrends[3].TypeField[7]=0;
	TTrends[3].NameField[7]="О2топкаП1";
	TTrends[3].AboutField[7]="О2топкаП1";
	TTrends[3].TypeField[8]=0;
	TTrends[3].NameField[8]="ПолГРК";
	TTrends[3].AboutField[8]="ПолГРК";
	TTrends[3].TypeField[9]=0;
	TTrends[3].NameField[9]="РвоздЗаВЗП_Л";
	TTrends[3].AboutField[9]="РвоздЗаВЗП_Л";
	TTrends[3].TypeField[10]=0;
	TTrends[3].NameField[10]="РвоздЗаВЗП_П";
	TTrends[3].AboutField[10]="РвоздЗаВЗП_П";
	TTrends[3].TypeField[11]=0;
	TTrends[3].NameField[11]="РвоздСредн_Д";
	TTrends[3].AboutField[11]="РвоздСредн_Д";
	TTrends[3].TypeField[12]=0;
	TTrends[3].NameField[12]="РгазДоОК";
	TTrends[3].AboutField[12]="РгазДоОК";
	TTrends[3].TypeField[13]=0;
	TTrends[3].NameField[13]="РгазЗаГРК";
	TTrends[3].AboutField[13]="РгазЗаГРК";
	TTrends[3].TypeField[14]=0;
	TTrends[3].NameField[14]="РпПарК8";
	TTrends[3].AboutField[14]="РпПарК8";
	TTrends[3].TypeField[15]=0;
	TTrends[3].NameField[15]="ТгазНаКотел";
	TTrends[3].AboutField[15]="ТгазНаКотел";
	TTrends[3].TypeField[16]=0;
	TTrends[3].NameField[16]="ТпарЗаКотл8";
	TTrends[3].AboutField[16]="ТпарЗаКотл8";
	TTrends[3].TypeField[17]=0;
	TTrends[3].NameField[17]="FнПродСум";
	TTrends[3].AboutField[17]="FнПродСум";
	TTrends[3].TypeField[18]=0;
	TTrends[3].NameField[18]="FпитВоды";
	TTrends[3].AboutField[18]="FпитВоды";
	TTrends[3].TypeField[19]=1;
	TTrends[3].NameField[19]="ЧасыРаботыК8";
	TTrends[3].AboutField[19]="ЧасыРаботыК8";
	TTrends[3].TypeField[20]=1;
	TTrends[3].NameField[20]="ЧасыРаботыК8М";
	TTrends[3].AboutField[20]="ЧасыРаботыК8М";
	TTrends[3].TypeField[21]=0;
	TTrends[3].NameField[21]="ПолПК255";
	TTrends[3].AboutField[21]="ПолПК255";
	TTrends[3].TypeField[22]=0;
	TTrends[3].NameField[22]="ПолРК255";
	TTrends[3].AboutField[22]="ПолРК255";
	TTrends[3].TypeField[23]=0;
	TTrends[3].NameField[23]="РвБарабане";
	TTrends[3].AboutField[23]="РвБарабане";
	TTrends[3].TypeField[24]=0;
	TTrends[3].NameField[24]="РпвДоСУП";
	TTrends[3].AboutField[24]="РпвДоСУП";
	TTrends[3].TypeField[25]=0;
	TTrends[3].NameField[25]="ТвздЗаВЗП2Л1";
	TTrends[3].AboutField[25]="ТвздЗаВЗП2Л1";
	TTrends[3].TypeField[26]=0;
	TTrends[3].NameField[26]="ТвздЗаВЗП2П1";
	TTrends[3].AboutField[26]="ТвздЗаВЗП2П1";
	TTrends[3].TypeField[27]=0;
	TTrends[3].NameField[27]="ТвздНарЛ1";
	TTrends[3].AboutField[27]="ТвздНарЛ1";
	TTrends[3].TypeField[28]=0;
	TTrends[3].NameField[28]="ТвздНарП1";
	TTrends[3].AboutField[28]="ТвздНарП1";
	TTrends[3].TypeField[29]=0;
	TTrends[3].NameField[29]="ТзаКалорЛ1";
	TTrends[3].AboutField[29]="ТзаКалорЛ1";
	TTrends[3].TypeField[30]=0;
	TTrends[3].NameField[30]="ТзаКалорП1";
	TTrends[3].AboutField[30]="ТзаКалорП1";
	TTrends[3].TypeField[31]=0;
	TTrends[3].NameField[31]="ТпитВоды";
	TTrends[3].AboutField[31]="ТпитВоды";
	TTrends[3].TypeField[32]=0;
	TTrends[3].NameField[32]="ТуходГазовЛ1";
	TTrends[3].AboutField[32]="ТуходГазовЛ1";
	TTrends[3].TypeField[33]=0;
	TTrends[3].NameField[33]="ТуходГазовП1";
	TTrends[3].AboutField[33]="ТуходГазовП1";
	TTrends[3].TypeField[34]=0;
	TTrends[3].NameField[34]="ТшкУСО";
	TTrends[3].AboutField[34]="ТшкУСО";

	TTrends[4].IdField="id";
	TTrends[4].DateField="date";

	TTrends[4].TypeField[0]=0;
	TTrends[4].NameField[0]="FгазКотел";
	TTrends[4].AboutField[0]="FгазКотел";
	TTrends[4].TypeField[1]=0;
	TTrends[4].NameField[1]="FпПар";
	TTrends[4].AboutField[1]="FпПар";
	TTrends[4].TypeField[2]=0;
	TTrends[4].NameField[2]="SвТопкеСредн";
	TTrends[4].AboutField[2]="SвТопкеСредн";
	TTrends[4].TypeField[3]=2;
	TTrends[4].NameField[3]="ЕстьФакел";
	TTrends[4].AboutField[3]="ЕстьФакел";
	TTrends[4].TypeField[4]=0;
	TTrends[4].NameField[4]="Н_Барабан";
	TTrends[4].AboutField[4]="Н_Барабан";
	TTrends[4].TypeField[5]=0;
	TTrends[4].NameField[5]="О2среднее";
	TTrends[4].AboutField[5]="О2среднее";
	TTrends[4].TypeField[6]=0;
	TTrends[4].NameField[6]="О2топкаЛ1";
	TTrends[4].AboutField[6]="О2топкаЛ1";
	TTrends[4].TypeField[7]=0;
	TTrends[4].NameField[7]="О2топкаП1";
	TTrends[4].AboutField[7]="О2топкаП1";
	TTrends[4].TypeField[8]=0;
	TTrends[4].NameField[8]="ПолГРК";
	TTrends[4].AboutField[8]="ПолГРК";
	TTrends[4].TypeField[9]=0;
	TTrends[4].NameField[9]="РвоздЗаВЗП_Л";
	TTrends[4].AboutField[9]="РвоздЗаВЗП_Л";
	TTrends[4].TypeField[10]=0;
	TTrends[4].NameField[10]="РвоздЗаВЗП_П";
	TTrends[4].AboutField[10]="РвоздЗаВЗП_П";
	TTrends[4].TypeField[11]=0;
	TTrends[4].NameField[11]="РвоздСредн_Д";
	TTrends[4].AboutField[11]="РвоздСредн_Д";
	TTrends[4].TypeField[12]=0;
	TTrends[4].NameField[12]="РгазДоОК";
	TTrends[4].AboutField[12]="РгазДоОК";
	TTrends[4].TypeField[13]=0;
	TTrends[4].NameField[13]="РгазЗаГРК";
	TTrends[4].AboutField[13]="РгазЗаГРК";
	TTrends[4].TypeField[14]=0;
	TTrends[4].NameField[14]="РпПарК9";
	TTrends[4].AboutField[14]="РпПарК9";
	TTrends[4].TypeField[15]=0;
	TTrends[4].NameField[15]="ТгазНаКотел";
	TTrends[4].AboutField[15]="ТгазНаКотел";
	TTrends[4].TypeField[16]=0;
	TTrends[4].NameField[16]="ТпарЗаКотл9";
	TTrends[4].AboutField[16]="ТпарЗаКотл9";
	TTrends[4].TypeField[17]=1;
	TTrends[4].NameField[17]="ЧасыРаботыК9";
	TTrends[4].AboutField[17]="ЧасыРаботыК9";
	TTrends[4].TypeField[18]=1;
	TTrends[4].NameField[18]="ЧасыРаботыК9М";
	TTrends[4].AboutField[18]="ЧасыРаботыК9М";
	TTrends[4].TypeField[19]=0;
	TTrends[4].NameField[19]="FнПродувки2";
	TTrends[4].AboutField[19]="FнПродувки2";
	TTrends[4].TypeField[20]=0;
	TTrends[4].NameField[20]="FпитВоды";
	TTrends[4].AboutField[20]="FпитВоды";
	TTrends[4].TypeField[21]=0;
	TTrends[4].NameField[21]="ПолПК255";
	TTrends[4].AboutField[21]="ПолПК255";
	TTrends[4].TypeField[22]=0;
	TTrends[4].NameField[22]="ПолРК255";
	TTrends[4].AboutField[22]="ПолРК255";
	TTrends[4].TypeField[23]=0;
	TTrends[4].NameField[23]="РвБарабане";
	TTrends[4].AboutField[23]="РвБарабане";
	TTrends[4].TypeField[24]=0;
	TTrends[4].NameField[24]="РпвДоСУП";
	TTrends[4].AboutField[24]="РпвДоСУП";
	TTrends[4].TypeField[25]=0;
	TTrends[4].NameField[25]="ТвздЗаВЗП2Л1";
	TTrends[4].AboutField[25]="ТвздЗаВЗП2Л1";
	TTrends[4].TypeField[26]=0;
	TTrends[4].NameField[26]="ТвздЗаВЗП2П1";
	TTrends[4].AboutField[26]="ТвздЗаВЗП2П1";
	TTrends[4].TypeField[27]=0;
	TTrends[4].NameField[27]="ТвздНарЛ1";
	TTrends[4].AboutField[27]="ТвздНарЛ1";
	TTrends[4].TypeField[28]=0;
	TTrends[4].NameField[28]="ТвздНарП1";
	TTrends[4].AboutField[28]="ТвздНарП1";
	TTrends[4].TypeField[29]=0;
	TTrends[4].NameField[29]="ТзаКалорЛ1";
	TTrends[4].AboutField[29]="ТзаКалорЛ1";
	TTrends[4].TypeField[30]=0;
	TTrends[4].NameField[30]="ТзаКалорП1";
	TTrends[4].AboutField[30]="ТзаКалорП1";
	TTrends[4].TypeField[31]=0;
	TTrends[4].NameField[31]="ТпитВоды";
	TTrends[4].AboutField[31]="ТпитВоды";
	TTrends[4].TypeField[32]=0;
	TTrends[4].NameField[32]="ТуходГазовЛ1";
	TTrends[4].AboutField[32]="ТуходГазовЛ1";
	TTrends[4].TypeField[33]=0;
	TTrends[4].NameField[33]="ТуходГазовП1";
	TTrends[4].AboutField[33]="ТуходГазовП1";
	TTrends[4].TypeField[34]=0;
	TTrends[4].NameField[34]="ТшкУСО";
	TTrends[4].AboutField[34]="ТшкУСО";

	TTrends[5].IdField="id";
	TTrends[5].DateField="date";

	TTrends[5].TypeField[0]=0;
	TTrends[5].NameField[0]="FгазКотел";
	TTrends[5].AboutField[0]="FгазКотел";
	TTrends[5].TypeField[1]=0;
	TTrends[5].NameField[1]="FпПар";
	TTrends[5].AboutField[1]="FпПар";
	TTrends[5].TypeField[2]=0;
	TTrends[5].NameField[2]="SвТопкеРег";
	TTrends[5].AboutField[2]="SвТопкеРег";
	TTrends[5].TypeField[3]=2;
	TTrends[5].NameField[3]="ЕстьФакел";
	TTrends[5].AboutField[3]="ЕстьФакел";
	TTrends[5].TypeField[4]=0;
	TTrends[5].NameField[4]="Н_Барабан";
	TTrends[5].AboutField[4]="Н_Барабан";
	TTrends[5].TypeField[5]=0;
	TTrends[5].NameField[5]="О2среднее";
	TTrends[5].AboutField[5]="О2среднее";
	TTrends[5].TypeField[6]=0;
	TTrends[5].NameField[6]="О2топкаЛ1";
	TTrends[5].AboutField[6]="О2топкаЛ1";
	TTrends[5].TypeField[7]=0;
	TTrends[5].NameField[7]="О2топкаП1";
	TTrends[5].AboutField[7]="О2топкаП1";
	TTrends[5].TypeField[8]=0;
	TTrends[5].NameField[8]="ПолГРК";
	TTrends[5].AboutField[8]="ПолГРК";
	TTrends[5].TypeField[9]=0;
	TTrends[5].NameField[9]="РвБарабане";
	TTrends[5].AboutField[9]="РвБарабане";
	TTrends[5].TypeField[10]=0;
	TTrends[5].NameField[10]="РвоздЛ";
	TTrends[5].AboutField[10]="РвоздЛ";
	TTrends[5].TypeField[11]=0;
	TTrends[5].NameField[11]="РвоздП";
	TTrends[5].AboutField[11]="РвоздП";
	TTrends[5].TypeField[12]=0;
	TTrends[5].NameField[12]="РгазДоОК";
	TTrends[5].AboutField[12]="РгазДоОК";
	TTrends[5].TypeField[13]=0;
	TTrends[5].NameField[13]="РгазЗаГРК";
	TTrends[5].AboutField[13]="РгазЗаГРК";
	TTrends[5].TypeField[14]=0;
	TTrends[5].NameField[14]="РпПарК10";
	TTrends[5].AboutField[14]="РпПарК10";
	TTrends[5].TypeField[15]=0;
	TTrends[5].NameField[15]="ТгазНаКотел";
	TTrends[5].AboutField[15]="ТгазНаКотел";
	TTrends[5].TypeField[16]=0;
	TTrends[5].NameField[16]="ТпарЗаКотл10";
	TTrends[5].AboutField[16]="ТпарЗаКотл10";
	TTrends[5].TypeField[17]=1;
	TTrends[5].NameField[17]="ЧислоГорелок";
	TTrends[5].AboutField[17]="ЧислоГорелок";

	TTrends[6].IdField="id";
	TTrends[6].DateField="date";

	TTrends[6].TypeField[0]=0;
	TTrends[6].NameField[0]="FпПарК3";
	TTrends[6].AboutField[0]="FпПарК3";
	TTrends[6].TypeField[1]=0;
	TTrends[6].NameField[1]="FпПарК4";
	TTrends[6].AboutField[1]="FпПарК4";
	TTrends[6].TypeField[2]=0;
	TTrends[6].NameField[2]="FпПарТ3К5пр";
	TTrends[6].AboutField[2]="FпПарТ3К5пр";
	TTrends[6].TypeField[3]=0;
	TTrends[6].NameField[3]="FпПарТ3К6пр";
	TTrends[6].AboutField[3]="FпПарТ3К6пр";
	TTrends[6].TypeField[4]=0;
	TTrends[6].NameField[4]="FпПарТ4К7пр";
	TTrends[6].AboutField[4]="FпПарТ4К7пр";
	TTrends[6].TypeField[5]=0;
	TTrends[6].NameField[5]="FпПарТ4К8пр";
	TTrends[6].AboutField[5]="FпПарТ4К8пр";
	TTrends[6].TypeField[6]=0;
	TTrends[6].NameField[6]="FпПарТ5К10пр";
	TTrends[6].AboutField[6]="FпПарТ5К10пр";
	TTrends[6].TypeField[7]=0;
	TTrends[6].NameField[7]="FпПарТ6К11пр";
	TTrends[6].AboutField[7]="FпПарТ6К11пр";
	TTrends[6].TypeField[8]=0;
	TTrends[6].NameField[8]="N_ТГ3";
	TTrends[6].AboutField[8]="N_ТГ3";
	TTrends[6].TypeField[9]=0;
	TTrends[6].NameField[9]="N_ТГ4";
	TTrends[6].AboutField[9]="N_ТГ4";
	TTrends[6].TypeField[10]=0;
	TTrends[6].NameField[10]="N_ТГ5";
	TTrends[6].AboutField[10]="N_ТГ5";
	TTrends[6].TypeField[11]=0;
	TTrends[6].NameField[11]="N_ТГ6";
	TTrends[6].AboutField[11]="N_ТГ6";
	TTrends[6].TypeField[12]=0;
	TTrends[6].NameField[12]="T_МПК";
	TTrends[6].AboutField[12]="T_МПК";
	TTrends[6].TypeField[13]=0;
	TTrends[6].NameField[13]="ЗдКПС2оч";
	TTrends[6].AboutField[13]="ЗдКПС2оч";
	TTrends[6].TypeField[14]=0;
	TTrends[6].NameField[14]="ЗдКПС3оч";
	TTrends[6].AboutField[14]="ЗдКПС3оч";
	TTrends[6].TypeField[15]=0;
	TTrends[6].NameField[15]="КПС_Рп2оч";
	TTrends[6].AboutField[15]="КПС_Рп2оч";
	TTrends[6].TypeField[16]=0;
	TTrends[6].NameField[16]="КПС_Рп2оч2";
	TTrends[6].AboutField[16]="КПС_Рп2оч2";
	TTrends[6].TypeField[17]=0;
	TTrends[6].NameField[17]="КПС_Рп2оч3";
	TTrends[6].AboutField[17]="КПС_Рп2оч3";
	TTrends[6].TypeField[18]=0;
	TTrends[6].NameField[18]="КПС_Рп3оч";
	TTrends[6].AboutField[18]="КПС_Рп3оч";
	TTrends[6].TypeField[19]=0;
	TTrends[6].NameField[19]="КПС_Рп3оч2";
	TTrends[6].AboutField[19]="КПС_Рп3оч2";
	TTrends[6].TypeField[20]=0;
	TTrends[6].NameField[20]="РпПарТ2К3";
	TTrends[6].AboutField[20]="РпПарТ2К3";
	TTrends[6].TypeField[21]=0;
	TTrends[6].NameField[21]="РпПарТ2К4";
	TTrends[6].AboutField[21]="РпПарТ2К4";
	TTrends[6].TypeField[22]=0;
	TTrends[6].NameField[22]="РпПарТ3К5";
	TTrends[6].AboutField[22]="РпПарТ3К5";
	TTrends[6].TypeField[23]=0;
	TTrends[6].NameField[23]="РпПарТ3К6";
	TTrends[6].AboutField[23]="РпПарТ3К6";
	TTrends[6].TypeField[24]=0;
	TTrends[6].NameField[24]="РпПарТ4К7";
	TTrends[6].AboutField[24]="РпПарТ4К7";
	TTrends[6].TypeField[25]=0;
	TTrends[6].NameField[25]="РпПарТ4К8";
	TTrends[6].AboutField[25]="РпПарТ4К8";
	TTrends[6].TypeField[26]=0;
	TTrends[6].NameField[26]="РпПарТ5К10";
	TTrends[6].AboutField[26]="РпПарТ5К10";
	TTrends[6].TypeField[27]=0;
	TTrends[6].NameField[27]="РпПарТ6К11";
	TTrends[6].AboutField[27]="РпПарТ6К11";
	TTrends[6].TypeField[28]=0;
	TTrends[6].NameField[28]="Т_усо4";
	TTrends[6].AboutField[28]="Т_усо4";
	TTrends[6].TypeField[29]=0;
	TTrends[6].NameField[29]="ТпЗаК10";
	TTrends[6].AboutField[29]="ТпЗаК10";
	TTrends[6].TypeField[30]=0;
	TTrends[6].NameField[30]="ТпарЗаК11";
	TTrends[6].AboutField[30]="ТпарЗаК11";
	TTrends[6].TypeField[31]=0;
	TTrends[6].NameField[31]="ТпарЗаК3";
	TTrends[6].AboutField[31]="ТпарЗаК3";
	TTrends[6].TypeField[32]=0;
	TTrends[6].NameField[32]="ТпарЗаК4";
	TTrends[6].AboutField[32]="ТпарЗаК4";
	TTrends[6].TypeField[33]=0;
	TTrends[6].NameField[33]="ТпарЗаКотл5";
	TTrends[6].AboutField[33]="ТпарЗаКотл5";
	TTrends[6].TypeField[34]=0;
	TTrends[6].NameField[34]="ТпарЗаКотл6";
	TTrends[6].AboutField[34]="ТпарЗаКотл6";
	TTrends[6].TypeField[35]=0;
	TTrends[6].NameField[35]="ТпарЗаКотл7";
	TTrends[6].AboutField[35]="ТпарЗаКотл7";
	TTrends[6].TypeField[36]=0;
	TTrends[6].NameField[36]="ТпарЗаКотл8";
	TTrends[6].AboutField[36]="ТпарЗаКотл8";
	TTrends[6].TypeField[37]=0;
	TTrends[6].NameField[37]="ТпарЗаКотл9";
	TTrends[6].AboutField[37]="ТпарЗаКотл9";
	TTrends[6].TypeField[38]=0;
	TTrends[6].NameField[38]="ТпарКПС2_1";
	TTrends[6].AboutField[38]="ТпарКПС2_1";
	TTrends[6].TypeField[39]=0;
	TTrends[6].NameField[39]="ТпарКПС2_10";
	TTrends[6].AboutField[39]="ТпарКПС2_10";
	TTrends[6].TypeField[40]=0;
	TTrends[6].NameField[40]="ТпарКПС2_11";
	TTrends[6].AboutField[40]="ТпарКПС2_11";
	TTrends[6].TypeField[41]=0;
	TTrends[6].NameField[41]="ТпарКПС2_2";
	TTrends[6].AboutField[41]="ТпарКПС2_2";
	TTrends[6].TypeField[42]=0;
	TTrends[6].NameField[42]="ТпарКПС2_3";
	TTrends[6].AboutField[42]="ТпарКПС2_3";
	TTrends[6].TypeField[43]=0;
	TTrends[6].NameField[43]="ТпарКПС2_4";
	TTrends[6].AboutField[43]="ТпарКПС2_4";
	TTrends[6].TypeField[44]=0;
	TTrends[6].NameField[44]="ТпарКПС2_5";
	TTrends[6].AboutField[44]="ТпарКПС2_5";
	TTrends[6].TypeField[45]=0;
	TTrends[6].NameField[45]="ТпарКПС2_6";
	TTrends[6].AboutField[45]="ТпарКПС2_6";
	TTrends[6].TypeField[46]=0;
	TTrends[6].NameField[46]="ТпарКПС2_7";
	TTrends[6].AboutField[46]="ТпарКПС2_7";
	TTrends[6].TypeField[47]=0;
	TTrends[6].NameField[47]="ТпарКПС2_8";
	TTrends[6].AboutField[47]="ТпарКПС2_8";
	TTrends[6].TypeField[48]=0;
	TTrends[6].NameField[48]="ТпарКПС2_9";
	TTrends[6].AboutField[48]="ТпарКПС2_9";
	TTrends[6].TypeField[49]=0;
	TTrends[6].NameField[49]="ТпарКПС3_12";
	TTrends[6].AboutField[49]="ТпарКПС3_12";
	TTrends[6].TypeField[50]=0;
	TTrends[6].NameField[50]="ТпарКПС3_13";
	TTrends[6].AboutField[50]="ТпарКПС3_13";
	TTrends[6].TypeField[51]=0;
	TTrends[6].NameField[51]="ТпарТ2К3";
	TTrends[6].AboutField[51]="ТпарТ2К3";
	TTrends[6].TypeField[52]=0;
	TTrends[6].NameField[52]="ТпарТ2К4";
	TTrends[6].AboutField[52]="ТпарТ2К4";
	TTrends[6].TypeField[53]=0;
	TTrends[6].NameField[53]="ТпарТ3К5";
	TTrends[6].AboutField[53]="ТпарТ3К5";
	TTrends[6].TypeField[54]=0;
	TTrends[6].NameField[54]="ТпарТ3К6";
	TTrends[6].AboutField[54]="ТпарТ3К6";
	TTrends[6].TypeField[55]=0;
	TTrends[6].NameField[55]="ТпарТ4К7";
	TTrends[6].AboutField[55]="ТпарТ4К7";
	TTrends[6].TypeField[56]=0;
	TTrends[6].NameField[56]="ТпарТ4К8";
	TTrends[6].AboutField[56]="ТпарТ4К8";
	TTrends[6].TypeField[57]=0;
	TTrends[6].NameField[57]="ТпарТ5К10";
	TTrends[6].AboutField[57]="ТпарТ5К10";
	TTrends[6].TypeField[58]=0;
	TTrends[6].NameField[58]="ТпарТ6К11";
	TTrends[6].AboutField[58]="ТпарТ6К11";

	int dn=6;
	if (WorkMode==2) dn=-1;

	TTrends[dn+1].IdField="id";
	TTrends[dn+1].DateField="date";

	TTrends[dn+1].TypeField[0]=0;
	TTrends[dn+1].NameField[0]="3_T.39a";
	TTrends[dn+1].AboutField[0]="Реактивная мощность";
	TTrends[dn+1].TypeField[1]=0;
	TTrends[dn+1].NameField[1]="3_T.40a";
	TTrends[dn+1].AboutField[1]="Iвозб.ротора";
	TTrends[dn+1].TypeField[2]=0;
	TTrends[dn+1].NameField[2]="3_T.41a";
	TTrends[dn+1].AboutField[2]="Тхол.газ.ген.т1";
	TTrends[dn+1].TypeField[3]=0;
	TTrends[dn+1].NameField[3]="3_T.42a";
	TTrends[dn+1].AboutField[3]="Тхол.газ.ген.т2";
	TTrends[dn+1].TypeField[4]=0;
	TTrends[dn+1].NameField[4]="3_T.43a";
	TTrends[dn+1].AboutField[4]="Тхол.газ.ген.т3";
	TTrends[dn+1].TypeField[5]=0;
	TTrends[dn+1].NameField[5]="3_T.44a";
	TTrends[dn+1].AboutField[5]="Тхол.газ.ген.т4";
	TTrends[dn+1].TypeField[6]=0;
	TTrends[dn+1].NameField[6]="3_T.45a";
	TTrends[dn+1].AboutField[6]="Твыхл.ч.слева";
	TTrends[dn+1].TypeField[7]=0;
	TTrends[dn+1].NameField[7]="3_T.46a";
	TTrends[dn+1].AboutField[7]="Твыхл.ч.справа";
	TTrends[dn+1].TypeField[8]=0;
	TTrends[dn+1].NameField[8]="3_T.47a";
	TTrends[dn+1].AboutField[8]="Тмасл.п.МО";
	TTrends[dn+1].TypeField[9]=0;
	TTrends[dn+1].NameField[9]="3_T.48a";
	TTrends[dn+1].AboutField[9]="Тмасл.н.сл.подш1";
	TTrends[dn+1].TypeField[10]=0;
	TTrends[dn+1].NameField[10]="3_T.49a";
	TTrends[dn+1].AboutField[10]="Тмасл.н.сл.подш2";
	TTrends[dn+1].TypeField[11]=0;
	TTrends[dn+1].NameField[11]="3_T.50a";
	TTrends[dn+1].AboutField[11]="Тмасл.н.сл.подш3";
	TTrends[dn+1].TypeField[12]=0;
	TTrends[dn+1].NameField[12]="3_T.51a";
	TTrends[dn+1].AboutField[12]="Тмасл.н.сл.подш4";
	TTrends[dn+1].TypeField[13]=0;
	TTrends[dn+1].NameField[13]="3_T.52a";
	TTrends[dn+1].AboutField[13]="Тмасл.н.сл.подш5";
	TTrends[dn+1].TypeField[14]=0;
	TTrends[dn+1].NameField[14]="3_T.53a";
	TTrends[dn+1].AboutField[14]="Тмасл.н.сл.подш6";
	TTrends[dn+1].TypeField[15]=0;
	TTrends[dn+1].NameField[15]="3_T.54a";
	TTrends[dn+1].AboutField[15]="Тмасл.н.сл.подш7";
	TTrends[dn+1].TypeField[16]=0;
	TTrends[dn+1].NameField[16]="3_T.55a";
	TTrends[dn+1].AboutField[16]="Тмасл.н.сл.подш8";
	TTrends[dn+1].TypeField[17]=0;
	TTrends[dn+1].NameField[17]="3_T.56a";
	TTrends[dn+1].AboutField[17]="Тбабб.оп.подш1";
	TTrends[dn+1].TypeField[18]=0;
	TTrends[dn+1].NameField[18]="3_T.57a";
	TTrends[dn+1].AboutField[18]="Тбабб.оп.подш2";
	TTrends[dn+1].TypeField[19]=0;
	TTrends[dn+1].NameField[19]="3_T.58a";
	TTrends[dn+1].AboutField[19]="Тбабб.оп.подш3";
	TTrends[dn+1].TypeField[20]=0;
	TTrends[dn+1].NameField[20]="3_T.59a";
	TTrends[dn+1].AboutField[20]="Тбабб.оп.подш4";
	TTrends[dn+1].TypeField[21]=0;
	TTrends[dn+1].NameField[21]="3_T.60a";
	TTrends[dn+1].AboutField[21]="Тбабб.оп.подш5";
	TTrends[dn+1].TypeField[22]=0;
	TTrends[dn+1].NameField[22]="3_T.61a";
	TTrends[dn+1].AboutField[22]="Тбабб.оп.подш6";
	TTrends[dn+1].TypeField[23]=0;
	TTrends[dn+1].NameField[23]="3_T.62a";
	TTrends[dn+1].AboutField[23]="Тбабб.оп.подш7";
	TTrends[dn+1].TypeField[24]=0;
	TTrends[dn+1].NameField[24]="3_T.63a";
	TTrends[dn+1].AboutField[24]="Тбабб.оп.подш8т1";
	TTrends[dn+1].TypeField[25]=0;
	TTrends[dn+1].NameField[25]="3_T.64a";
	TTrends[dn+1].AboutField[25]="Тбабб.оп.подш8т2";
	TTrends[dn+1].TypeField[26]=0;
	TTrends[dn+1].NameField[26]="3_T.65a";
	TTrends[dn+1].AboutField[26]="Траб.кол.1.ОУП";
	TTrends[dn+1].TypeField[27]=0;
	TTrends[dn+1].NameField[27]="3_T.66a";
	TTrends[dn+1].AboutField[27]="Траб.кол.2.ОУП";
	TTrends[dn+1].TypeField[28]=0;
	TTrends[dn+1].NameField[28]="3_T.67a";
	TTrends[dn+1].AboutField[28]="Траб.кол.3.ОУП";
	TTrends[dn+1].TypeField[29]=0;
	TTrends[dn+1].NameField[29]="3_T.68a";
	TTrends[dn+1].AboutField[29]="Траб.кол.4.ОУП";
	TTrends[dn+1].TypeField[30]=0;
	TTrends[dn+1].NameField[30]="3_T.70a";
	TTrends[dn+1].AboutField[30]="Траб.кол.5.ОУП";
	TTrends[dn+1].TypeField[31]=0;
	TTrends[dn+1].NameField[31]="3_T.71a";
	TTrends[dn+1].AboutField[31]="Траб.кол.6.ОУП";
	TTrends[dn+1].TypeField[32]=0;
	TTrends[dn+1].NameField[32]="3_T.72a";
	TTrends[dn+1].AboutField[32]="Траб.кол.7.ОУП";
	TTrends[dn+1].TypeField[33]=0;
	TTrends[dn+1].NameField[33]="3_T.73a";
	TTrends[dn+1].AboutField[33]="Траб.кол.8.ОУП";
	TTrends[dn+1].TypeField[34]=0;
	TTrends[dn+1].NameField[34]="3_T.75a";
	TTrends[dn+1].AboutField[34]="Туст.кол.1.ОУП";
	TTrends[dn+1].TypeField[35]=0;
	TTrends[dn+1].NameField[35]="3_T.76a";
	TTrends[dn+1].AboutField[35]="Туст.кол.2.ОУП";
	TTrends[dn+1].TypeField[36]=0;
	TTrends[dn+1].NameField[36]="3_T.77a";
	TTrends[dn+1].AboutField[36]="Туст.кол.3.ОУП";
	TTrends[dn+1].TypeField[37]=0;
	TTrends[dn+1].NameField[37]="3_T.78a";
	TTrends[dn+1].AboutField[37]="Туст.кол.4.ОУП";
	TTrends[dn+1].TypeField[38]=0;
	TTrends[dn+1].NameField[38]="3_T.81a";
	TTrends[dn+1].AboutField[38]="Туст.кол.5.ОУП";
	TTrends[dn+1].TypeField[39]=0;
	TTrends[dn+1].NameField[39]="3_T.82a";
	TTrends[dn+1].AboutField[39]="Туст.кол.6.ОУП";
	TTrends[dn+1].TypeField[40]=0;
	TTrends[dn+1].NameField[40]="3_T.83a";
	TTrends[dn+1].AboutField[40]="Туст.кол.7.ОУП";
	TTrends[dn+1].TypeField[41]=0;
	TTrends[dn+1].NameField[41]="3_T.84a";
	TTrends[dn+1].AboutField[41]="Туст.кол.8.ОУП";
	TTrends[dn+1].TypeField[42]=0;
	TTrends[dn+1].NameField[42]="3_T.87a";
	TTrends[dn+1].AboutField[42]="Тпара уплотнения";
	TTrends[dn+1].TypeField[43]=0;
	TTrends[dn+1].NameField[43]="3_T.88a";
	TTrends[dn+1].AboutField[43]="ТЦВДз.ст.верх";
	TTrends[dn+1].TypeField[44]=0;
	TTrends[dn+1].NameField[44]="3_T.89a";
	TTrends[dn+1].AboutField[44]="ТЦВДз.ст.низ";
	TTrends[dn+1].TypeField[45]=0;
	TTrends[dn+1].NameField[45]="3_T.90a";
	TTrends[dn+1].AboutField[45]="Тфл.ЦВДза1ст.сл.верх";
	TTrends[dn+1].TypeField[46]=0;
	TTrends[dn+1].NameField[46]="3_T.91a";
	TTrends[dn+1].AboutField[46]="Тфл.ЦВДза1ст.сп.верх";
	TTrends[dn+1].TypeField[47]=0;
	TTrends[dn+1].NameField[47]="3_T.92a";
	TTrends[dn+1].AboutField[47]="Т6шп.ЦВДслева";
	TTrends[dn+1].TypeField[48]=0;
	TTrends[dn+1].NameField[48]="3_T.93a";
	TTrends[dn+1].AboutField[48]="Т6шп.ЦВДсправа";
	TTrends[dn+1].TypeField[49]=0;
	TTrends[dn+1].NameField[49]="3_T.94a";
	TTrends[dn+1].AboutField[49]="ТЦВДза6ст.верх";
	 TTrends[dn+1].TypeField[50]=0;
	 TTrends[dn+1].NameField[50]="3_T.95a";
	TTrends[dn+1].AboutField[50]="Тфл.ЦВДза6ст.сл.верх";
	 TTrends[dn+1].TypeField[51]=0;
	 TTrends[dn+1].NameField[51]="3_T.96a";
	TTrends[dn+1].AboutField[51]="Тфл.ЦВДза6ст.сп.верх";
	 TTrends[dn+1].TypeField[52]=0;
	 TTrends[dn+1].NameField[52]="3_T.97a";
	TTrends[dn+1].AboutField[52]="Т11шп.ЦВДсправа";
	 TTrends[dn+1].TypeField[53]=0;
	 TTrends[dn+1].NameField[53]="3_T.98a";
	TTrends[dn+1].AboutField[53]="Тпара за 1й ст.";
	 TTrends[dn+1].TypeField[54]=0;
	 TTrends[dn+1].NameField[54]="3_T.110_1";
	TTrends[dn+1].AboutField[54]="Т оп.подш. >80";
	 TTrends[dn+1].TypeField[55]=0;
	 TTrends[dn+1].NameField[55]="3_T.110_2";
	TTrends[dn+1].AboutField[55]="Т хол.газа >42";
	 TTrends[dn+1].TypeField[56]=0;
	 TTrends[dn+1].NameField[56]="3_T.110_3";
	TTrends[dn+1].AboutField[56]="Т кол.ОУП >80";
	 TTrends[dn+1].TypeField[57]=0;
	 TTrends[dn+1].NameField[57]="3_T.110_4";
	TTrends[dn+1].AboutField[57]="Тмасла на сл.подш.>65";
	 TTrends[dn+1].TypeField[58]=0;
	 TTrends[dn+1].NameField[58]="3_T.99a";
	TTrends[dn+1].AboutField[58]="Т ЦВД за 6ст.низ";
	 TTrends[dn+1].TypeField[59]=0;
	 TTrends[dn+1].NameField[59]="3_T.100a";
	TTrends[dn+1].AboutField[59]="Вакуум в конд.";
	 TTrends[dn+1].TypeField[60]=0;
	 TTrends[dn+1].NameField[60]="3_T.69a";
	TTrends[dn+1].AboutField[60]="пустой";
	 TTrends[dn+1].TypeField[61]=0;
	 TTrends[dn+1].NameField[61]="3_T.74a";
	TTrends[dn+1].AboutField[61]="пустой";
	 TTrends[dn+1].TypeField[62]=0;
	 TTrends[dn+1].NameField[62]="3_T.79a";
	TTrends[dn+1].AboutField[62]="пустой";
	 TTrends[dn+1].TypeField[63]=0;
	 TTrends[dn+1].NameField[63]="3_T.80a";
	TTrends[dn+1].AboutField[63]="пустой";
	 TTrends[dn+1].TypeField[64]=0;
	 TTrends[dn+1].NameField[64]="3_T.85a";
	TTrends[dn+1].AboutField[64]="пустой";
	 TTrends[dn+1].TypeField[65]=0;
	 TTrends[dn+1].NameField[65]="3_T.86a";
	TTrends[dn+1].AboutField[65]="пустой";

	TTrends[dn+2].IdField="id";
	TTrends[dn+2].DateField="date";

	 TTrends[dn+2].TypeField[0]=0;
	 TTrends[dn+2].NameField[0]="3_Tah1_F";
	TTrends[dn+2].AboutField[0]="Тахометр";
	 TTrends[dn+2].TypeField[1]=0;
	 TTrends[dn+2].NameField[1]="3_АктМощн";
	TTrends[dn+2].AboutField[1]="Активная мощность";
	 TTrends[dn+2].TypeField[2]=0;
	 TTrends[dn+2].NameField[2]="3_IV1_S";
	TTrends[dn+2].AboutField[2]="Искривление вала";
	 TTrends[dn+2].TypeField[3]=0;
	 TTrends[dn+2].NameField[3]="3_OS1_S";
	TTrends[dn+2].AboutField[3]="Осевой сдвиг 1";
	 TTrends[dn+2].TypeField[4]=0;
	 TTrends[dn+2].NameField[4]="3_OS2_S";
	TTrends[dn+2].AboutField[4]="Осевой сдвиг 2";
	 TTrends[dn+2].TypeField[5]=0;
	 TTrends[dn+2].NameField[5]="3_OS3_S";
	TTrends[dn+2].AboutField[5]="Осевой сдвиг 3";
	 TTrends[dn+2].TypeField[6]=0;
	 TTrends[dn+2].NameField[6]="3_ARV_S";
	TTrends[dn+2].AboutField[6]="Абс.расш.верхн.";
	 TTrends[dn+2].TypeField[7]=0;
	 TTrends[dn+2].NameField[7]="3_ARC_S";
	TTrends[dn+2].AboutField[7]="Абс.расш.нижн.";
	 TTrends[dn+2].TypeField[8]=0;
	 TTrends[dn+2].NameField[8]="3_ORV_S";
	TTrends[dn+2].AboutField[8]="Отн.расш.верхн.";
	 TTrends[dn+2].TypeField[9]=0;
	 TTrends[dn+2].NameField[9]="3_ORC_S";
	TTrends[dn+2].AboutField[9]="Отн.расш.средн.";
	 TTrends[dn+2].TypeField[10]=0;
	 TTrends[dn+2].NameField[10]="3_ORN_S";
	TTrends[dn+2].AboutField[10]="Отн.расш.нижн.";

	 TTrends[dn+2].TypeField[11]=0;
	 TTrends[dn+2].NameField[11]="3_Op1O_Ve";
	TTrends[dn+2].AboutField[11]="Подш1,Осев,Ve";
	 TTrends[dn+2].TypeField[12]=0;
	 TTrends[dn+2].NameField[12]="3_Op1P_Ve";
	TTrends[dn+2].AboutField[12]="Подш1,Попер,Ve";
	 TTrends[dn+2].TypeField[13]=0;
	 TTrends[dn+2].NameField[13]="3_Op1V_Ve";
	TTrends[dn+2].AboutField[13]="Подш1,Верт,Ve";
//	 TTrends[dn+2].TypeField[14]=0;
//	 TTrends[dn+2].NameField[14]="3_Op1O_S";
//	TTrends[dn+2].AboutField[14]="Подш1,Осев,S";
//	 TTrends[dn+2].TypeField[15]=0;
//	 TTrends[dn+2].NameField[15]="3_Op1P_S";
//	TTrends[dn+2].AboutField[15]="Подш1,Попер,S";
//	 TTrends[dn+2].TypeField[16]=0;
//	 TTrends[dn+2].NameField[16]="3_Op1V_S";
//	TTrends[dn+2].AboutField[16]="Подш1,Верт,S";

	 TTrends[dn+2].TypeField[14]=0;
	 TTrends[dn+2].NameField[14]="3_Op2O_Ve";
	TTrends[dn+2].AboutField[14]="Подш2,Осев,Ve";
	 TTrends[dn+2].TypeField[15]=0;
	 TTrends[dn+2].NameField[15]="3_Op2P_Ve";
	TTrends[dn+2].AboutField[15]="Подш2,Попер,Ve";
	 TTrends[dn+2].TypeField[16]=0;
	 TTrends[dn+2].NameField[16]="3_Op2V_Ve";
	TTrends[dn+2].AboutField[16]="Подш2,Верт,Ve";
//	 TTrends[dn+2].TypeField[20]=0;
//	 TTrends[dn+2].NameField[20]="3_Op2O_S";
//	TTrends[dn+2].AboutField[20]="Подш2,Осев,S";
//	 TTrends[dn+2].TypeField[21]=0;
//	 TTrends[dn+2].NameField[21]="3_Op2P_S";
//	TTrends[dn+2].AboutField[21]="Подш2,Попер,S";
//	 TTrends[dn+2].TypeField[22]=0;
//	 TTrends[dn+2].NameField[22]="3_Op2V_S";
//	TTrends[dn+2].AboutField[22]="Подш2,Верт,S";

	 TTrends[dn+2].TypeField[17]=0;
	 TTrends[dn+2].NameField[17]="3_Op3O_Ve";
	TTrends[dn+2].AboutField[17]="Подш3,Осев,Ve";
	 TTrends[dn+2].TypeField[18]=0;
	 TTrends[dn+2].NameField[18]="3_Op3P_Ve";
	TTrends[dn+2].AboutField[18]="Подш3,Попер,Ve";
	 TTrends[dn+2].TypeField[19]=0;
	 TTrends[dn+2].NameField[19]="3_Op3V_Ve";
	TTrends[dn+2].AboutField[19]="Подш3,Верт,Ve";
//	 TTrends[dn+2].TypeField[26]=0;
//	 TTrends[dn+2].NameField[26]="3_Op3O_S";
//	TTrends[dn+2].AboutField[26]="Подш3,Осев,S";
//	 TTrends[dn+2].TypeField[27]=0;
//	 TTrends[dn+2].NameField[27]="3_Op3P_S";
//	TTrends[dn+2].AboutField[27]="Подш3,Попер,S";
//	 TTrends[dn+2].TypeField[28]=0;
//	 TTrends[dn+2].NameField[28]="3_Op3V_S";
//	TTrends[dn+2].AboutField[28]="Подш3,Верт,S";

	 TTrends[dn+2].TypeField[20]=0;
	 TTrends[dn+2].NameField[20]="3_Op4O_Ve";
	TTrends[dn+2].AboutField[20]="Подш4,Осев,Ve";
	 TTrends[dn+2].TypeField[21]=0;
	 TTrends[dn+2].NameField[21]="3_Op4P_Ve";
	TTrends[dn+2].AboutField[21]="Подш4,Попер,Ve";
	 TTrends[dn+2].TypeField[22]=0;
	 TTrends[dn+2].NameField[22]="3_Op4V_Ve";
	TTrends[dn+2].AboutField[22]="Подш4,Верт,Ve";
//	 TTrends[dn+2].TypeField[32]=0;
//	 TTrends[dn+2].NameField[32]="3_Op4O_S";
//	TTrends[dn+2].AboutField[32]="Подш4,Осев,S";
//	 TTrends[dn+2].TypeField[33]=0;
//	 TTrends[dn+2].NameField[33]="3_Op4P_S";
//	TTrends[dn+2].AboutField[33]="Подш4,Попер,S";
//	 TTrends[dn+2].TypeField[34]=0;
//	 TTrends[dn+2].NameField[34]="3_Op4V_S";
//	TTrends[dn+2].AboutField[34]="Подш4,Верт,S";

	 TTrends[dn+2].TypeField[23]=0;
	 TTrends[dn+2].NameField[23]="3_Op5O_Ve";
	TTrends[dn+2].AboutField[23]="Подш5,Осев,Ve";
	 TTrends[dn+2].TypeField[24]=0;
	 TTrends[dn+2].NameField[24]="3_Op5P_Ve";
	TTrends[dn+2].AboutField[24]="Подш5,Попер,Ve";
	 TTrends[dn+2].TypeField[25]=0;
	 TTrends[dn+2].NameField[25]="3_Op5V_Ve";
	TTrends[dn+2].AboutField[25]="Подш5,Верт,Ve";
//	 TTrends[dn+2].TypeField[38]=0;
//	 TTrends[dn+2].NameField[38]="3_Op5O_S";
//	TTrends[dn+2].AboutField[38]="Подш5,Осев,S";
//	 TTrends[dn+2].TypeField[39]=0;
//	 TTrends[dn+2].NameField[39]="3_Op5P_S";
//	TTrends[dn+2].AboutField[39]="Подш5,Попер,S";
//	 TTrends[dn+2].TypeField[40]=0;
//	 TTrends[dn+2].NameField[40]="3_Op5V_S";
//	TTrends[dn+2].AboutField[40]="Подш5,Верт,S";

	 TTrends[dn+2].TypeField[26]=0;
	 TTrends[dn+2].NameField[26]="3_Op6O_Ve";
	TTrends[dn+2].AboutField[26]="Подш6,Осев,Ve";
	 TTrends[dn+2].TypeField[27]=0;
	 TTrends[dn+2].NameField[27]="3_Op6P_Ve";
	TTrends[dn+2].AboutField[27]="Подш6,Попер,Ve";
	 TTrends[dn+2].TypeField[28]=0;
	 TTrends[dn+2].NameField[28]="3_Op6V_Ve";
	TTrends[dn+2].AboutField[28]="Подш6,Верт,Ve";
//	 TTrends[dn+2].TypeField[44]=0;
//	 TTrends[dn+2].NameField[44]="3_Op6O_S";
//	TTrends[dn+2].AboutField[44]="Подш6,Осев,S";
//	 TTrends[dn+2].TypeField[45]=0;
//	 TTrends[dn+2].NameField[45]="3_Op6P_S";
//	TTrends[dn+2].AboutField[45]="Подш6,Попер,S";
//	 TTrends[dn+2].TypeField[46]=0;
//	 TTrends[dn+2].NameField[46]="3_Op6V_S";
//	TTrends[dn+2].AboutField[46]="Подш6,Верт,S";

	 TTrends[dn+2].TypeField[29]=0;
	 TTrends[dn+2].NameField[29]="3_Op7O_Ve";
	TTrends[dn+2].AboutField[29]="Подш7,Осев,Ve";
	 TTrends[dn+2].TypeField[30]=0;
	 TTrends[dn+2].NameField[30]="3_Op7P_Ve";
	TTrends[dn+2].AboutField[30]="Подш7,Попер,Ve";
	 TTrends[dn+2].TypeField[31]=0;
	 TTrends[dn+2].NameField[31]="3_Op7V_Ve";
	TTrends[dn+2].AboutField[31]="Подш7,Верт,Ve";
//	 TTrends[dn+2].TypeField[50]=0;
//	 TTrends[dn+2].NameField[50]="3_Op7O_S";
//	TTrends[dn+2].AboutField[50]="Подш7,Осев,S";
//	 TTrends[dn+2].TypeField[51]=0;
//	 TTrends[dn+2].NameField[51]="3_Op7P_S";
//	TTrends[dn+2].AboutField[51]="Подш7,Попер,S";
//	 TTrends[dn+2].TypeField[52]=0;
//	 TTrends[dn+2].NameField[52]="3_Op7V_S";
//	TTrends[dn+2].AboutField[52]="Подш7,Верт,S";

	 TTrends[dn+2].TypeField[32]=0;
	 TTrends[dn+2].NameField[32]="3_Op8O_Ve";
	TTrends[dn+2].AboutField[32]="Подш8,Осев,Ve";
	 TTrends[dn+2].TypeField[33]=0;
	 TTrends[dn+2].NameField[33]="3_Op8P_Ve";
	TTrends[dn+2].AboutField[33]="Подш8,Попер,Ve";
	 TTrends[dn+2].TypeField[34]=0;
	 TTrends[dn+2].NameField[34]="3_Op8V_Ve";
	TTrends[dn+2].AboutField[34]="Подш8,Верт,Ve";
//	 TTrends[dn+2].TypeField[56]=0;
//	 TTrends[dn+2].NameField[56]="3_Op8O_S";
//	TTrends[dn+2].AboutField[56]="Подш8,Осев,S";
//	 TTrends[dn+2].TypeField[57]=0;
//	 TTrends[dn+2].NameField[57]="3_Op8P_S";
//	TTrends[dn+2].AboutField[57]="Подш8,Попер,S";
//	 TTrends[dn+2].TypeField[58]=0;
//	 TTrends[dn+2].NameField[58]="3_Op8V_S";
//	TTrends[dn+2].AboutField[58]="Подш8,Верт,S";


	 TTrends[dn+2].TypeField[35]=0;
	 TTrends[dn+2].NameField[35]="Reg___";
	TTrends[dn+2].AboutField[35]="Reg___";
	 TTrends[dn+2].TypeField[36]=0;
	 TTrends[dn+2].NameField[36]="Reg1__";
	TTrends[dn+2].AboutField[36]="Reg1__";
	 TTrends[dn+2].TypeField[37]=0;
	 TTrends[dn+2].NameField[37]="Reg2__";
	TTrends[dn+2].AboutField[37]="Reg2__";



	TTrends[dn+3].IdField="id";
	TTrends[dn+3].DateField="date";

	 TTrends[dn+3].TypeField[0]=0;
	 TTrends[dn+3].NameField[0]="6_Op1V_Ve";
	TTrends[dn+3].AboutField[0]="Подш1,Верт,Ve";
	 TTrends[dn+3].TypeField[1]=0;
	 TTrends[dn+3].NameField[1]="6_Op1P_Ve";
	TTrends[dn+3].AboutField[1]="Подш1,Попер,Ve";
	 TTrends[dn+3].TypeField[2]=0;
	 TTrends[dn+3].NameField[2]="6_Op1O_Ve";
	TTrends[dn+3].AboutField[2]="Подш1,Осев,Ve";
//	 TTrends[dn+3].TypeField[3]=0;
//	 TTrends[dn+3].NameField[3]="6_Op1V_S";
//	TTrends[dn+3].AboutField[3]="Подш1,Верт,S";
//	 TTrends[dn+3].TypeField[4]=0;
//	 TTrends[dn+3].NameField[4]="6_Op1P_S";
//	TTrends[dn+3].AboutField[4]="Подш1,Попер,S";
//	 TTrends[dn+3].TypeField[5]=0;
//	 TTrends[dn+3].NameField[5]="6_Op1O_S";
//	TTrends[dn+3].AboutField[5]="Подш1,Осев,S";
	 TTrends[dn+3].TypeField[3]=0;
	 TTrends[dn+3].NameField[3]="6_Op2V_Ve";
	TTrends[dn+3].AboutField[3]="Подш2,Верт,Ve";
	 TTrends[dn+3].TypeField[4]=0;
	 TTrends[dn+3].NameField[4]="6_Op2P_Ve";
	TTrends[dn+3].AboutField[4]="Подш2,Попер,Ve";
	 TTrends[dn+3].TypeField[5]=0;
	 TTrends[dn+3].NameField[5]="6_Op2O_Ve";
	TTrends[dn+3].AboutField[5]="Подш2,Осев,Ve";
//	 TTrends[dn+3].TypeField[9]=0;
//	 TTrends[dn+3].NameField[9]="6_Op2V_S";
//	TTrends[dn+3].AboutField[9]="Подш2,Верт,S";
//	 TTrends[dn+3].TypeField[10]=0;
//	 TTrends[dn+3].NameField[10]="6_Op2P_S";
//	TTrends[dn+3].AboutField[10]="Подш2,Попер,S";
//	 TTrends[dn+3].TypeField[11]=0;
//	 TTrends[dn+3].NameField[11]="6_Op2O_S";
//	TTrends[dn+3].AboutField[11]="Подш2,Осев,S";
	 TTrends[dn+3].TypeField[6]=0;
	 TTrends[dn+3].NameField[6]="6_Op3V_Ve";
	TTrends[dn+3].AboutField[6]="Подш3,Верт,Ve";
	 TTrends[dn+3].TypeField[7]=0;
	 TTrends[dn+3].NameField[7]="6_Op3P_Ve";
	TTrends[dn+3].AboutField[7]="Подш3,Попер,Ve";
	 TTrends[dn+3].TypeField[8]=0;
	 TTrends[dn+3].NameField[8]="6_Op3O_Ve";
	TTrends[dn+3].AboutField[8]="Подш3,Осев,Ve";
//	 TTrends[dn+3].TypeField[15]=0;
//	 TTrends[dn+3].NameField[15]="6_Op3V_S";
//	TTrends[dn+3].AboutField[15]="Подш3,Верт,S";
//	 TTrends[dn+3].TypeField[16]=0;
//	 TTrends[dn+3].NameField[16]="6_Op3P_S";
//  TTrends[dn+3].AboutField[16]="Подш3,Попер,S";
//	 TTrends[dn+3].TypeField[17]=0;
//	 TTrends[dn+3].NameField[17]="6_Op3O_S";
//	TTrends[dn+3].AboutField[17]="Подш3,Осев,S";
	 TTrends[dn+3].TypeField[9]=0;
	 TTrends[dn+3].NameField[9]="6_Op4V_Ve";
	TTrends[dn+3].AboutField[9]="Подш4,Верт,Ve";
	 TTrends[dn+3].TypeField[10]=0;
	 TTrends[dn+3].NameField[10]="6_Op4P_Ve";
	TTrends[dn+3].AboutField[10]="Подш4,Попер,Ve";
	 TTrends[dn+3].TypeField[11]=0;
	 TTrends[dn+3].NameField[11]="6_Op4O_Ve";
	TTrends[dn+3].AboutField[11]="Подш4,Осев,Ve";
//	 TTrends[dn+3].TypeField[12]=0;
//	 TTrends[dn+3].NameField[21]="6_Op4V_S";
//	TTrends[dn+3].AboutField[21]="Подш4,Верт,S";
//	 TTrends[dn+3].TypeField[22]=0;
//	 TTrends[dn+3].NameField[22]="6_Op4P_S";
//	TTrends[dn+3].AboutField[22]="Подш4,Попер,S";
//	 TTrends[dn+3].TypeField[23]=0;
//	 TTrends[dn+3].NameField[23]="6_Op4O_S";
//	TTrends[dn+3].AboutField[23]="Подш4,Осев,S";
	 TTrends[dn+3].TypeField[12]=0;
	 TTrends[dn+3].NameField[12]="6_Op5V_Ve";
	TTrends[dn+3].AboutField[12]="Подш5,Верт,Ve";
	 TTrends[dn+3].TypeField[13]=0;
	 TTrends[dn+3].NameField[13]="6_Op5P_Ve";
	TTrends[dn+3].AboutField[13]="Подш5,Попер,Ve";
	 TTrends[dn+3].TypeField[14]=0;
	 TTrends[dn+3].NameField[14]="6_Op5O_Ve";
	TTrends[dn+3].AboutField[14]="Подш5,Осев,Ve";
//	 TTrends[dn+3].TypeField[27]=0;
//	 TTrends[dn+3].NameField[27]="6_Op5V_S";
//	TTrends[dn+3].AboutField[27]="Подш5,Верт,S";
//	 TTrends[dn+3].TypeField[28]=0;
//	 TTrends[dn+3].NameField[28]="6_Op5P_S";
//	TTrends[dn+3].AboutField[28]="Подш5,Попер,S";
//	 TTrends[dn+3].TypeField[29]=0;
//	 TTrends[dn+3].NameField[29]="6_Op5O_S";
//	TTrends[dn+3].AboutField[29]="Подш5,Осев,S";
	 TTrends[dn+3].TypeField[15]=0;
	 TTrends[dn+3].NameField[15]="6_Op6V_Ve";
	TTrends[dn+3].AboutField[15]="Подш6,Верт,Ve";
	 TTrends[dn+3].TypeField[16]=0;
	 TTrends[dn+3].NameField[16]="6_Op6P_Ve";
	TTrends[dn+3].AboutField[16]="Подш6,Попер,Ve";
	 TTrends[dn+3].TypeField[17]=0;
	 TTrends[dn+3].NameField[17]="6_Op6O_Ve";
	TTrends[dn+3].AboutField[17]="Подш6,Осев,Ve";
//	 TTrends[dn+3].TypeField[33]=0;
//	 TTrends[dn+3].NameField[33]="6_Op6V_S";
//	TTrends[dn+3].AboutField[33]="Подш6,Верт,S";
//	 TTrends[dn+3].TypeField[34]=0;
//	 TTrends[dn+3].NameField[34]="6_Op6P_S";
//	TTrends[dn+3].AboutField[34]="Подш6,Попер,S";
//	 TTrends[dn+3].TypeField[35]=0;
//	 TTrends[dn+3].NameField[35]="6_Op6O_S";
//	TTrends[dn+3].AboutField[35]="Подш6,Осев,S";
	 TTrends[dn+3].TypeField[18]=0;
	 TTrends[dn+3].NameField[18]="6_OS1_S";
	TTrends[dn+3].AboutField[18]="ОС1";
	 TTrends[dn+3].TypeField[19]=0;
	 TTrends[dn+3].NameField[19]="6_OS2_S";
	TTrends[dn+3].AboutField[19]="ОС2";
	 TTrends[dn+3].TypeField[20]=0;
	 TTrends[dn+3].NameField[20]="6_OS3_S";
	TTrends[dn+3].AboutField[20]="ОС3";
	 TTrends[dn+3].TypeField[21]=0;
	 TTrends[dn+3].NameField[21]="6_ORV1_1_S";
	TTrends[dn+3].AboutField[21]="ОР РВД";
	 TTrends[dn+3].TypeField[22]=0;
	 TTrends[dn+3].NameField[22]="6_ORN1_1_S";
	TTrends[dn+3].AboutField[22]="ОР РНД";
	 TTrends[dn+3].TypeField[23]=0;
	 TTrends[dn+3].NameField[23]="6_ARV1_1_S";
	TTrends[dn+3].AboutField[23]="АРЦВД прав.";
	 TTrends[dn+3].TypeField[24]=0;
	 TTrends[dn+3].NameField[24]="6_ARN1_1_S";
	TTrends[dn+3].AboutField[24]="АРЦНД прав.";
	 TTrends[dn+3].TypeField[25]=0;
	 TTrends[dn+3].NameField[25]="6_Мощн_Акт";
	TTrends[dn+3].AboutField[25]="Активная мощность";
	 TTrends[dn+3].TypeField[26]=0;
	 TTrends[dn+3].NameField[26]="6_Мощн_Реакт";
	TTrends[dn+3].AboutField[26]="Реактивная мощность";
	 TTrends[dn+3].TypeField[27]=0;
	 TTrends[dn+3].NameField[27]="6_Ротор_Ток";
	TTrends[dn+3].AboutField[27]="Ток ротора";
	 TTrends[dn+3].TypeField[28]=0;
	 TTrends[dn+3].NameField[28]="6_IV1_S";
	TTrends[dn+3].AboutField[28]="ИВ1 искривление";
	 TTrends[dn+3].TypeField[29]=0;
	 TTrends[dn+3].NameField[29]="6_Tah1_F";
	TTrends[dn+3].AboutField[29]="Тахометр 1";

	for (int i=0;i<NumListTrends;i++)
	{
		for (int j=0;j<ListTrends[i].NumVar;j++)
		{
			AnsiString st="param";
			wchar_t s[4];
			_itow (j, s, 10);
			st.operator +=(s);
			TTrends[i].NameFieldDB[j]=st;
		}
		TrendsDS[i] = new TADOQuery(this);
		TablesDS[i] = new TADOQuery(this);
		TrendsConn[i] = new TADOConnection(this);
		ParamsDS[i] = new TADOQuery(this);
		ValuesDS[i] = new TADOQuery(this);
		DBSQLConn[i] = new TADOConnection(this);
//		TrendsConn2[i] = new SqlConnection(this);
		ConnectDB(i);
//		ReOpenDB(i, true);
	}

	Paused=false;

	if (StringGrid1->RowCount>0)
	{
//		StringGrid1->CellRect(1,1)
		StringGrid1->Row=0;
		StringGrid1->Col=1;
//		StringGrid1SelectCell(TForm1, 1, 0, false);
	}
	else
	SelectedOPCServer=-1;
	TrendsStoredRemoute=true;
	TrendsStoredLocal=true;
	TimeReOpen=(double)((double)(18)/(double)(24));
	NumDayTrendFile=(double)(2);

///////////////////////////////
///////////////////////////////Сервер(
///////////////////////////////
	// CLSID_Svr identifies this Server
	// (Created with guidgen.exe)
	// minimum server refresh rate is 250 ms
//	InitWTOPCsvr ((BYTE *)&CLSID_Svr,250);

///	EnableWriteNotification (&WriteNotifyProc, TRUE);
///	EnableDisconnectNotification (&DisconnectProc);
///	EnableDeviceRead (&DeviceReadProc);

	//
	// If you want to override the default A&E operation of WtOPCsvr.dll
	// define your own callback object and overload the functions of interest.
//	pCallback = new (CMyAECallback);
//	SetAEServerCallback (TRUE, pCallback);

///////////////////////////////
///////////////////////////////)Сервер
///////////////////////////////

	LogMessage1=L"Запуск программы";
	LogMessage2=L"инициализация произведена";
	Log();

	Timer2->Enabled=true;
//	startform->Close();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::ReOpenDB(int nTrendFile, bool SaveOld)
{
	TrendsDS[nTrendFile]->Close();
	TablesDS[nTrendFile]->Close();
	TrendsConn[nTrendFile]->Close();

	char* SrcFilePath=FullNullTrendFilePath.c_str();
	char* DestFilePath=FullTrendFilePath[nTrendFile].c_str();
	if (SaveOld)
	{
		LogMessage1=L"Работа с БД";
		LogMessage2=L"Сохранение БД";
		Log();
		AnsiString NewFullTrendFilePath=ExtractFilePath(Application->ExeName);
		NewFullTrendFilePath.operator +=(TrendPath[nTrendFile]);
		NewFullTrendFilePath.operator +=(TrendFile[nTrendFile]);
		NewFullTrendFilePath.operator +=("_");
		TDateTime dt=Now();
		NewFullTrendFilePath.operator +=(dt.CurrentDate().DateString());
		wchar_t s[16];
		_itow (dt.CurrentTime().Val, s, 10);
		NewFullTrendFilePath.operator +=("_");
		NewFullTrendFilePath.operator +=(dt.CurrentTime().Val);
		NewFullTrendFilePath.operator +=(".sbs");
		if (!MoveFile(DestFilePath, NewFullTrendFilePath.c_str()))
		{
			DWORD err=GetLastError();
			wchar_t sterr[5];
			_itow (err, sterr, 10);
			String stmess="Не удалось сохранить текущий файл трендов ";
			stmess.operator +=(DestFilePath);
//			Application->MessageBox((LPCWSTR)stmess.c_str() , sterr, MB_OK);
			LogMessage1=L"Работа с БД";
			LogMessage2=L"Не удалось сохранить текущий файл трендов";
			Log();
		}
	}
	if (!CopyFile(SrcFilePath, DestFilePath, false))
	{
		DWORD err=GetLastError();
		wchar_t sterr[5];
		_itow (err, sterr, 10);
		String stmess="Не удалось скопировать базовый файл трендов ";
		stmess.operator +=(FullNullTrendFilePath);
		stmess.operator +=("  в  ");
		stmess.operator +=(DestFilePath);
//		Application->MessageBox((LPCWSTR)stmess.c_str() , sterr, MB_OK);
		LogMessage1=L"Работа с БД";
		LogMessage2=L"Не удалось начать новый текущий файл трендов";
		Log();
	}
	ConnectDB(nTrendFile);
}
//---------------------------------------------------------------------------
void __fastcall TForm1::ReconnectDBSQL(int nDB)
{
	int idTrend=Form1->OPCServers[nDB].idTrend;
	int NT=0;
	AnsiString QueryTrend = "select * from tables";
//	Form1->TablesDS[idTrend]->Open();
	Form1->TablesDS[idTrend]->SQL->Clear();
	Form1->TablesDS[idTrend]->SQL->Add(QueryTrend);
//	Form1->TablesDS[idTrend]->ExecSQL();
	Form1->TablesDS[idTrend]->Open();
	Form1->TablesDS[idTrend]->First();
	while (!Form1->TablesDS[idTrend]->Eof)
	{   int NT1=Form1->TablesDS[idTrend]->FieldByName("valtable")->AsString.ToInt();
		if (NT1>NT) NT=NT1;
		Form1->TablesDS[idTrend]->Next();
	}
//	Form1->TablesDS[idTrend]->Close();
	NT++;
	TableVals[idTrend]="vals";
	wchar_t NTS[5];
	_itow (NT, NTS, 10);
	TableVals[idTrend]+=NTS;

//	AnsiString QueryTable="CREATE TABLE "+TableVals[idTrend]+" (id bigint NOT NULL, date datetime NULL, param int NULL, val real NULL)";
//	AnsiString QueryTable="CREATE TABLE "+TableVals[idTrend]+" (id bigint NOT NULL PRIMARY KEY CLUSTERED, date datetime NULL, param int NULL, val real NULL)";
	AnsiString QueryTable="CREATE TABLE "+TableVals[idTrend]+" (date datetime NULL, param int NULL, val real NULL)";
	if (!Form1->DBSQLConn[idTrend]->Connected)
	{
		Form1->DBSQLConn[idTrend]->Open();
		Form1->ValuesDS[idTrend]->Connection = Form1->DBSQLConn[idTrend];
	}
//	Form1->ValuesDS[idTrend]->Close();
	Form1->ValuesDS[idTrend]->SQL->Clear();
	Form1->ValuesDS[idTrend]->SQL->Add(QueryTable);
	Form1->ValuesDS[idTrend]->ExecSQL();
//	Form1->ValuesDS[idTrend]->Open();
	QueryTable="alter table "+TableVals[idTrend]+" add id int identity(1, 1)";//CREATE UNIQUE CLUSTERED INDEX idx1 ON "+TableVals[idTrend]+"(id)";//CREATE INDEX PrimaryKey ON "+TableVals[idTrend]+" (id) WITH PRIMARY";
	Form1->ValuesDS[idTrend]->SQL->Clear();
	Form1->ValuesDS[idTrend]->SQL->Add(QueryTable);
	Form1->ValuesDS[idTrend]->ExecSQL();
//	Form1->ValuesDS[idTrend]->Open();

	Form1->TablesDS[idTrend]->Last();
	if (Form1->TablesDS[idTrend]->RecordCount>0)
	{
//		Form1->TablesDS[idTrend]->Open();
		Form1->TablesDS[idTrend]->Edit();
		Form1->TablesDS[idTrend]->FieldByName("endtime")->AsDateTime = Now();
		Form1->TablesDS[idTrend]->Post();
//		Form1->TablesDS[idTrend]->Close();
	}
//	Form1->TablesDS[idTrend]->Open();
//	Form1->TablesDS[idTrend]->Last();
	Form1->TablesDS[idTrend]->Insert();
	Form1->TablesDS[idTrend]->FieldByName("starttime")->AsDateTime = Now();
	Form1->TablesDS[idTrend]->FieldByName("endtime")->AsDateTime = Now();
	Form1->TablesDS[idTrend]->FieldByName("paramtable")->AsString="params";
	Form1->TablesDS[idTrend]->FieldByName("valtable")->AsString=NTS;
	Form1->TablesDS[idTrend]->FieldByName("reccount")->AsInteger=0;
	Form1->TablesDS[idTrend]->Post();
	idTables[idTrend]=Form1->TablesDS[idTrend]->FieldByName("id")->AsInteger;
//	Form1->TablesDS[idTrend]->Close();


	QueryTable = "select * from "+TableVals[idTrend];
	Form1->ValuesDS[idTrend]->SQL->Clear();
	Form1->ValuesDS[idTrend]->SQL->Add(QueryTable);
//	Form1->ValuesDS[idTrend]->ExecSQL();// ->Open();
	Form1->ValuesDS[idTrend]->Open();
	Form1->ValuesDS[idTrend]->First();
}
//---------------------------------------------------------------------------
void __fastcall TForm1::ConnectDBSQL(int nDB)
{
	int idTrend=Form1->OPCServers[nDB].idTrend;
	String UserName = Form1->Edit1->Text;
	String PassWord = Form1->MaskEdit1->Text;
//	String Server = Form1->ComboBox4->Text+"\\ASUTP_OPCSERVER";
	String Server = Form1->ComboBox4->Text+"\\OPCSERVERASUTP1";
	try
	{
		String ConnString1 ="Provider=SQLOLEDB.1;";
		ConnString1 +="Data Source=%s;Initial Catalog=%s;User Id=%s;Password=%s;";
		String BD = Form1->TrendFile[idTrend];
		ConnString1 = Format(ConnString1, ARRAYOFCONST((Server, BD, UserName, PassWord)));
		if (Form1->DBSQLConn[idTrend]->Connected)
		{
			Form1->ParamsDS[idTrend]->Close();
			Form1->ValuesDS[idTrend]->Close();
			Form1->TablesDS[idTrend]->Close();
			Form1->DBSQLConn[idTrend]->Close();
		}
		Form1->DBSQLConn[idTrend]->ConnectionString = ConnString1;
		Form1->DBSQLConn[idTrend]->Open();
		Form1->TablesDS[idTrend]->Connection = Form1->DBSQLConn[idTrend];
		Form1->ParamsDS[idTrend]->Connection = Form1->DBSQLConn[idTrend];
		Form1->ValuesDS[idTrend]->Connection = Form1->DBSQLConn[idTrend];

		AnsiString QueryTrend = "select * from tables";
		Form1->TablesDS[idTrend]->SQL->Clear();
		Form1->TablesDS[idTrend]->SQL->Add(QueryTrend);
		Form1->TablesDS[idTrend]->Open();
		Form1->TablesDS[idTrend]->First();
		int NT=0, idT=0;
		while (!Form1->TablesDS[idTrend]->Eof)
		{
			if (Form1->TablesDS[idTrend]->FieldByName("reccount")->AsInteger<Form1->RecordsInTable)
			{
				NT=Form1->TablesDS[idTrend]->FieldByName("valtable")->AsString.ToInt();
				idT=Form1->TablesDS[idTrend]->FieldByName("id")->AsInteger;
			}
			Form1->TablesDS[idTrend]->Next();
		}
		if (NT==0)
		{
			ReconnectDBSQL(idTrend);
		}
		else
		{
			TableVals[idTrend]="vals";
			wchar_t NTS[5];
			_itow (NT, NTS, 10);
			TableVals[idTrend]+=NTS;
			idTables[idTrend]=idT;
		}

		QueryTrend = "select * from params";
		Form1->ParamsDS[idTrend]->SQL->Clear();
		Form1->ParamsDS[idTrend]->SQL->Add(QueryTrend);
		Form1->ParamsDS[idTrend]->Open();
		Form1->ParamsDS[idTrend]->First();

 /*		AnsiString TableVal="vals";
		int NT=0;
		while (!Form1->ParamsDS[idTrend]->Eof)
		{
			int MaxRecInTable=12*Form1->ParamsDS[idTrend]->RecordCount;
//			int MaxRecInTable=17280*Form1->ParamsDS[idTrend]->RecordCount;
			if (Form1->TablesDS[idTrend]->FieldByName("reccount")->AsInteger<MaxRecInTable)
			{
				break;
			}
			NT++;
			Form1->ParamsDS[idTrend]->Next();
		}
		wchar_t NTS[5];
		_itow (NT, NTS, 10);
		TableVal+=NTS;     */

		bool MustInit=false;
		int n=0;
		while (!Form1->ParamsDS[idTrend]->Eof)
		{

			if (Form1->ParamsDS[idTrend]->FieldByName("opcname")->AsAnsiString.AnsiCompareIC(OPCServers[idTrend].AboutOPCItem[n]))
			{
				MustInit=true;
				break;
			}
			n++;
			Form1->ParamsDS[idTrend]->Next();
		}
		if (MustInit | (Form1->ParamsDS[idTrend]->RecordCount==0))
		{
			InitDBSQL(idTrend);
		}
		Form1->ParamsDS[idTrend]->Close();
	} catch (...) {
	}
}
//---------------------------------------------------------------------------
void __fastcall TForm1::InitDBSQL(int nDB)
{
	int idTrend=Form1->OPCServers[nDB].idTrend;
	String UserName = Form1->Edit1->Text;
	String PassWord = Form1->MaskEdit1->Text;
//	String Server = Form1->ComboBox4->Text+"\\ASUTP_OPCSERVER";
	String Server = Form1->ComboBox4->Text+"\\OPCSERVERASUTP1";
	try
	{
		if (!Form1->DBSQLConn[idTrend]->Connected)
		{
			String ConnString1 ="Provider=SQLOLEDB.1;";
			ConnString1 +="Data Source=%s;Initial Catalog=%s;User Id=%s;Password=%s;";
			String BD = Form1->TrendFile[idTrend];
			ConnString1 = Format(ConnString1, ARRAYOFCONST((Server, BD, UserName, PassWord)));
			Form1->DBSQLConn[idTrend]->ConnectionString = ConnString1;
			Form1->ValuesDS[idTrend]->Connection = Form1->DBSQLConn[idTrend];
			Form1->ParamsDS[idTrend]->Connection = Form1->DBSQLConn[idTrend];
		}
		AnsiString QueryTrend = "select * from params";
		Form1->ParamsDS[idTrend]->SQL->Clear();
		Form1->ParamsDS[idTrend]->SQL->Add(QueryTrend);
		Form1->ParamsDS[idTrend]->Open();
		Form1->ParamsDS[idTrend]->First();
		while (!Form1->ParamsDS[idTrend]->Eof)
		{
			Form1->ParamsDS[idTrend]->Delete();
			int n;
			n=Form1->ParamsDS[idTrend]->RecordCount;
		}
		for (int i = 0; i < OPCServers[idTrend].nItems; i++)
		{
//Form1->ParamsDS[idTrend]->Edit()
Form1->ParamsDS[idTrend]->AppendRecord(
//Form1->ParamsDS[idTrend]->InsertRecord(
	ARRAYOFCONST((i,OPCServers[idTrend].AboutOPCItem[i], OPCServers[idTrend].AboutOPCItem[i])));
//	ARRAYOFCONST((NULL, OPCServers[idTrend].NameDBItem[i], OPCServers[idTrend].AboutOPCItem[i])));

//			Form1->ParamsDS1[idTrend]->FieldByName("name")=OPCServers[idTrend].AboutOPCItem[i];
//			Form1->ParamsDS1[idTrend]->FieldByName("opcname")=OPCServers[idTrend].NameDBItem[i];
		}
		Form1->ParamsDS[idTrend]->Close();

	} catch (...) {
	}
}
//---------------------------------------------------------------------------

void __fastcall TForm1::ConnectDB(int nTrend)
{
	bool Restart=false;

	String pr,dpr;
	pr="MSDataShape.1";
	dpr="Microsoft.Jet.OLEDB.4.0";
	const String ConnStr = "Provider=%s;Data Provider=%s;Data Source=%s";

	if (!FindTrendFile[nTrend])
	{
		FindTrendFile[nTrend]=true;
		char* SrcFilePath=FullNullTrendFilePath.c_str();
		char* DestFilePath=FullTrendFilePath[nTrend].c_str();
		CreateDirectory(FullTrendPath[nTrend].c_str(), NULL);
		if (!CopyFile(SrcFilePath, DestFilePath, false))
		{
			DWORD err=GetLastError();
			wchar_t sterr[5];
			_itow (err, sterr, 10);
			String stmess="Не удалось скопировать базовый файл трендов ";
			stmess.operator +=(FullNullTrendFilePath);
			stmess.operator +=("  в  ");
			stmess.operator +=(DestFilePath);
//			Application->MessageBox((LPCWSTR)stmess.c_str() , sterr, MB_OK);
			LogMessage1=L"Работа с БД";
			LogMessage2=L"Не удалось начать новый текущий файл трендов";
			Log();
		}
		Restart=true;
	}

	if (Restart)
	{
		LogMessage1=L"Работа с БД";
		LogMessage2=L"Для нормальной работы программы необходимо запустить ее еще раз";
		Log();
		LogMessage1=L"Работа с БД";
		LogMessage2=L"Перезапуск программы";
		Log();
		Application->MessageBox(L"Для нормальной работы программы необходимо запустить ее еще раз.", L"Инициализация", MB_OK);
		Application->Terminate();
	}
	else
	{
		AnsiString QueryTrend;

		String ConnString = Format (ConnStr, ARRAYOFCONST((pr,dpr,FullTrendFilePath[nTrend])));
		TrendsConn[nTrend]->ConnectionString=ConnString;
		TrendsConn[nTrend]->LoginPrompt = False;

		TrendsDS[nTrend]->Connection = TrendsConn[nTrend];
		TablesDS[nTrend]->Connection = TrendsConn[nTrend];

		QueryTrend = "select * from ";
		QueryTrend.operator +=(ListTrends[nTrend].TableName);
		QueryTrend.operator +=(ListTrends[nTrend].TableNamePrefix[0]);
		TrendsDS[nTrend]->SQL->Clear();
		TrendsDS[nTrend]->SQL->Add(QueryTrend);

		QueryTrend = "select * from tables";
		TablesDS[nTrend]->SQL->Clear();
		TablesDS[nTrend]->SQL->Add(QueryTrend);

		TrendsDS[nTrend]->Open();
		TablesDS[nTrend]->Open();
		int m;
		m=0;
	}
}
//---------------------------------------------------------------------------

void __fastcall TForm1::Button2Click(TObject *Sender)
{
	try
	{
        int ConnectResult=0;
        ConnectResult=ConnectToServer(SelectedOPCServer);
        StringGrid1->Repaint();
        if (ConnectResult<0)
		{
			wchar_t s[2];
			_itow (ConnectResult, s, 10);
//			MessageBox(Sender,
			Application->MessageBox(L"Не удалось соединиться с сервером.", s, MB_OK);
			SetButtonStatus(false);
		}
        else
		{
//            if (GroupAdd)//(0 != hOPCGroup)
//            {
                SetButtonStatus(true);
//                GetItems();
/*            }
            else
            {
                SetButtonStatus(false);
            }*/
        }
    }
    catch(...)
    {
		Application->MessageBox(L"Error", L"Error", MB_OK);
        Disconnect(-1);
		SetButtonStatus(false);
        throw;
    }
}
//---------------------------------------------------------------------------
int __fastcall TForm1::ConnectToServer(int NumOPCServer)
{
	User1 = Edit1->Text.c_str();
	Password1 = MaskEdit1->Text.c_str();
	Domain1 = "";//ComboBox2->Text.c_str();
	ServerName = ComboBox1->Text.c_str();


	USES_CONVERSION;
	COAUTHIDENTITY caId;
	caId.User = (unsigned short*) User1.m_str;
	caId.Password = (unsigned short*) Password1.m_str;
	caId.Domain = (unsigned short*) Domain1.m_str;
	caId.Flags = SEC_WINNT_AUTH_IDENTITY_UNICODE;//SEC_WINNT_AUTH_IDENTITY_ANSI;//
	int l1;
	int l2;
	int l3;
	l1=-1;
	l2=-2;
	l3=-3;
	l1=sizeof(caId.User);// User1.Length();
	l2=Password1.Length();
	l3=Domain1.Length();
	caId.UserLength = User1.Length();
	caId.PasswordLength = Password1.Length();
	caId.DomainLength = Domain1.Length();

	COAUTHINFO caInfo;
	caInfo.dwAuthnLevel = RPC_C_AUTHN_LEVEL_CONNECT;//RPC_C_AUTHN_LEVEL_DEFAULT;//RPC_C_AUTHN_LEVEL_PKT_PRIVACY,//
	caInfo.dwAuthnSvc = RPC_C_AUTHN_WINNT;//RPC_C_AUTHN_DEFAULT;//RPC_C_AUTHN_NONE;//RPC_C_AUTHN_GSS_KERBEROS;//
	caInfo.dwAuthzSvc = RPC_C_AUTHZ_NONE;//RPC_C_AUTHZ_DEFAULT;//RPC_C_AUTHZ_NAME;//
	caInfo.dwCapabilities = EOAC_NONE;//EOAC_DEFAULT;//EOAC_ANY_AUTHORITY;//
	caInfo.dwImpersonationLevel = RPC_C_IMP_LEVEL_IMPERSONATE;//RPC_C_IMP_LEVEL_DEFAULT;//RPC_C_IMP_LEVEL_IDENTIFY;//
	caInfo. pAuthIdentityData = &caId;
	caInfo.pwszServerPrincName = NULL;//WideString(ComboBox1->Text.c_str());//COLE_DEFAULT_PRINCIPAL;//

	COSERVERINFO csInfo;
	csInfo.pwszName = ServerName;
    csInfo.pAuthInfo = &caInfo;
    csInfo.dwReserved1 = 0;
    csInfo.dwReserved2 = 0;

    MULTI_QI mQI;
    mQI.pIID=&IID_IOPCServer;//&IID_IUnknown;//
    mQI.pItf=NULL;
    mQI.hr=S_OK;

    CLSID clsid;
    HRESULT hResult;
    CComBSTR clsidstr;
	hResult = CLSIDFromString(CLSIDOPCServer, &clsid);

	if (FAILED(hResult))
	{
		ServerConnect=false;
        return -1;
	}


	CoInitializeSecurity(
						NULL,
						RPC_C_AUTHN_WINNT,//-1,//
						NULL,
						NULL,
						RPC_C_AUTHN_LEVEL_CONNECT,//RPC_C_AUTHN_LEVEL_NONE,//
						RPC_C_IMP_LEVEL_IMPERSONATE,//RPC_C_IMP_LEVEL_IDENTIFY,//RPC_C_IMP_LEVEL_IDENTIFY,
						NULL,
						EOAC_NONE,
						NULL);

	hResult = CoCreateInstanceEx(clsid, NULL, CLSCTX_ALL,//CLSCTX_REMOTE_SERVER,//
								&csInfo, 1, &mQI);
    if ((FAILED(hResult)))
    {
        ServerConnect=false;
        if (hResult==E_INVALIDARG) return -11;
        if (hResult==REGDB_E_CLASSNOTREG) return -12;
		if (hResult==CLASS_E_NOAGGREGATION) return -13;
        if (hResult==CO_S_NOTALLINTERFACES) return -14;
        if (hResult==E_NOINTERFACE) return -15;
        if (hResult==E_ACCESSDENIED)  return -16;
        return -17;
	}
    else
    {
        HRESULT hr;
        hr = CoSetProxyBlanket(
                mQI.pItf,
                RPC_C_AUTHN_NONE,//RPC_C_AUTHN_WINNT,//
                RPC_C_AUTHZ_NONE,//RPC_C_AUTHZ_NAME,//
                NULL,//WideString(ComboBox1->Text.c_str()),//
                RPC_C_AUTHN_LEVEL_CONNECT,//RPC_C_AUTHN_LEVEL_PKT_PRIVACY,//RPC_C_AUTHN_LEVEL_DEFAULT,//
                RPC_C_IMP_LEVEL_IDENTIFY,//RPC_C_IMP_LEVEL_IMPERSONATE,//
                &caId,
                EOAC_NONE);//EOAC_ANY_AUTHORITY);//
        if (mQI.hr == S_OK) OPCServer =(IOPCServer*)mQI.pItf;
        ServerConnect=true;

        IClientSecurity *pcs = 0;
        hr = OPCServer->QueryInterface(IID_IClientSecurity, (void**)&pcs);
        if (SUCCEEDED(hr))
        {
            pcs->Release();
        }
	}

    if (OPCServer)
    {
        hResult = OPCServer->QueryInterface(IID_IOPCBrowseServerAddressSpace,
                                        reinterpret_cast<LPVOID*>(&OPCBrowseServerAdressSpace));
    	if (FAILED(hResult))
    	{
            BrowseServer=false;
            Disconnect(-1);
        if (hResult==E_INVALIDARG) return -21;
        if (hResult==REGDB_E_CLASSNOTREG) return -22;
        if (hResult==CLASS_E_NOAGGREGATION) return -23;
        if (hResult==CO_S_NOTALLINTERFACES) return -24;
		if (hResult==E_NOINTERFACE) return -25;
        if (hResult==E_ACCESSDENIED)  return -26;
            return -27;
        }
        else
        {
            BrowseServer=true;
        }
	}
    else
    {
        Disconnect(-1);
        return -20;
	}

		LPOLESTR GroupID = NULL;
        GUID GUIDGroupName;
        StringFromCLSID(clsid, &GroupID);
        DWORD dwRevisedRate = 0;
        hResult = OPCServer->AddGroup(GroupID, TRUE, m_dwRate, 0, 0, &m_fDeadBand,
                                        0, &hOPCGroup, &dwRevisedRate, IID_IUnknown,
                                        reinterpret_cast<LPUNKNOWN*>(&OPCGroup));
    	if (FAILED(hResult))
    	{
            GroupAdd=false;
            return -30;
        }
        else
        {
            GroupAdd=true;
        }
        m_dwRate = dwRevisedRate;
/*
        IConnectionPointContainer* ipCPC = NULL;
        ipCP = NULL;
        hResult = OPCGroup->QueryInterface(IID_IConnectionPointContainer,
                                            (void**)&ipCPC);
    	if (FAILED(hResult))
    	{
//            Application->MessageBox("IConnectionPointContainer не поддерживается сервером", "Сообщение", MB_OK);
        }
		else
        {
//            Application->MessageBox("IConnectionPointContainer поддерживается сервером", "Сообщение", MB_OK);
        	hResult = ipCPC->FindConnectionPoint(IID_IOPCDataCallback, &ipCP);
        	ipCPC->Release();
        	if (FAILED(hResult))
        	{
                ConnectPoint=false;
//                Application->MessageBox("OPCDataCallback не поддерживается сервером", "Сообщение", MB_OK);
        	}
            else
            {
//                Application->MessageBox("OPCDataCallback поддерживается сервером", "Сообщение", MB_OK);
                ConnectPoint=true;
                ipCallback = new Callback();
                dwAdvise = NULL;
            	hResult = ipCP->Advise(ipCallback, &dwAdvise);
            	if (FAILED(hResult))
            	{
                    ConnectPoint=false;
                    OnChange=false;
            		if (ipCP) ipCP->Release();
            		if (ipCallback) ipCallback->Release();
//                    Application->MessageBox("ipCP не поддерживается сервером", "Сообщение", MB_OK);
            	}
                else
                {
//                    Application->MessageBox("ipCP поддерживается сервером", "Сообщение", MB_OK);
                    OnChange=true;
                    ipCallback->EvDataChange = OnDataChange;
                    ipCallback->EvReadComplete  = OnReadComplete;
                }
            }
    	}
*/
        if (OPCGroup)
        {
            hResult = OPCGroup->QueryInterface(IID_IOPCSyncIO,
                                            reinterpret_cast<LPVOID*>(&OPCSyncIO));
            if (FAILED(hResult))
        	{
				SyncRead=false;
//                Application->MessageBox("IOPCSyncIO не поддерживается сервером", "Сообщение", MB_OK);
        	}
            else
            {
                SyncRead=true;
//                Application->MessageBox("IOPCSyncIO поддерживается сервером", "Сообщение", MB_OK);
            }
/*
			hResult = OPCGroup->QueryInterface(IID_IOPCAsyncIO2,
                                            reinterpret_cast<LPVOID*>(&OPCAsyncIO));
            if (FAILED(hResult))
        	{
                AsyncRead=false;
//                Application->MessageBox("IOPCAsyncIO2 не поддерживается сервером", "Сообщение", MB_OK);
			}
            else
            {
                AsyncRead=true;
//                Application->MessageBox("IOPCAsyncIO2 поддерживается сервером", "Сообщение", MB_OK);
			}
*/
		}
        CoTaskMemFree(GroupID);
        if (GroupID == NULL)
        {
			Application->MessageBox(L"All OK, but GroupID == NULL", L"Error", MB_OK);
			CoTaskMemFree(GroupID);
        }
		CLSIDOPCServerCurrent=CLSIDOPCServer;

//		CLSIDOPCServerCurrent.Empty();
//		CLSIDOPCServerCurrent.Attach(CLSIDOPCServer.m_str);

		NameOPCServerCurrent=NameOPCServer;
//		NameOPCServerCurrent.Empty();
//		NameOPCServerCurrent.Attach(NameOPCServer.m_str);

			OPCServers[SelectedOPCServer].CLSIDOPCServer=CLSIDOPCServer.m_str;
//			OPCServers[SelectedOPCServer].CLSIDOPCServer.Empty();
//			OPCServers[SelectedOPCServer].CLSIDOPCServer.Attach(CLSIDOPCServer.m_str);
			OPCServers[SelectedOPCServer].NameOPCServer=NameOPCServer.m_str;
//			OPCServers[SelectedOPCServer].NameOPCServer.Empty();
//			OPCServers[SelectedOPCServer].NameOPCServer.Attach(NameOPCServer.m_str);
			OPCServers[SelectedOPCServer].OPCServer = OPCServer;
			OPCServers[SelectedOPCServer].OPCBrowseServerAdressSpace=OPCBrowseServerAdressSpace;
			OPCServers[SelectedOPCServer].BrowseServer=BrowseServer;
			OPCServers[SelectedOPCServer].OPCGroup=OPCGroup;
            OPCServers[SelectedOPCServer].GroupAdd=GroupAdd;
            OPCServers[SelectedOPCServer].m_dwRate=m_dwRate;
            OPCServers[SelectedOPCServer].m_fDeadBand=m_fDeadBand;
            OPCServers[SelectedOPCServer].ipCP=ipCP;
            OPCServers[SelectedOPCServer].ConnectPoint=ConnectPoint;
            OPCServers[SelectedOPCServer].ipCallback=ipCallback;
            OPCServers[SelectedOPCServer].dwAdvise=dwAdvise;
            OPCServers[SelectedOPCServer].OnChange=OnChange;
            OPCServers[SelectedOPCServer].OPCSyncIO=OPCSyncIO;
            OPCServers[SelectedOPCServer].SyncRead=SyncRead;
			OPCServers[SelectedOPCServer].OPCAsyncIO=OPCAsyncIO;
            OPCServers[SelectedOPCServer].AsyncRead=AsyncRead;
            OPCServers[SelectedOPCServer].nItems=0;
			OPCServers[SelectedOPCServer].ServerConnect=true;
			GetItems(SelectedOPCServer);
			ConnectToItems(SelectedOPCServer);
			ListItems(SelectedOPCServer);

    Disconnect(-1);
    return 1;
}

//---------------------------------------------------------------------------
void __fastcall TForm1::ReleaseItem(int NumOPCServer)
{
	if (NumOPCServer>=0)
	{
        HRESULT hResult;
		if (0 != OPCServers[NumOPCServer].hOPCItem)
        {
            CComPtr<IOPCItemMgt> OPCItemMgt;
            if (OPCServers[NumOPCServer].OPCGroup)
            {
                hResult = OPCServers[NumOPCServer].OPCGroup->QueryInterface(IID_IOPCItemMgt,
									  reinterpret_cast<LPVOID*>(&OPCItemMgt));
                if (SUCCEEDED(hResult))
                {
                    HRESULT* phResult = NULL;
                    for (int i=0;i<OPCServers[NumOPCServer].nItems;i++)
                    {
                        hResult = OPCItemMgt->RemoveItems(1, &OPCServers[NumOPCServer].hOPCItem[i], &phResult);
                        if (SUCCEEDED(hResult))
                        {
                            HRESULT hr = phResult[0];
                            m_ptrMalloc->Free(phResult);
                            hResult = (HRESULT)hr;
                        }
                    }
                    OPCServers[NumOPCServer].nItems=0;
                }
            }
        }
	}
	else
	{
        HRESULT hResult;
        if (0 != hOPCItem)
        {
            CComPtr<IOPCItemMgt> OPCItemMgt;
            if (OPCGroup)
            {
                hResult = OPCGroup->QueryInterface(IID_IOPCItemMgt,
                                      reinterpret_cast<LPVOID*>(&OPCItemMgt));
                if (SUCCEEDED(hResult))
                {
                    HRESULT* phResult = NULL;
                    for (int i=0;i<nItems;i++)
                    {
                        hResult = OPCItemMgt->RemoveItems(1, &hOPCItem[i], &phResult);
                        if (SUCCEEDED(hResult))
                        {
                            HRESULT hr = phResult[0];
                            m_ptrMalloc->Free(phResult);
							hResult = (HRESULT)hr;
                        }
					}
					nItems=0;
                }
            }
        }
    }
}
//---------------------------------------------------------------------------
void __fastcall TForm1::Disconnect(int NumOPCServer)
{
    if (NumOPCServer>=0)
    {
        //OPCServers[NumOPCServer].CLSIDOPCServer.Empty();
        //OPCServers[NumOPCServer].NameOPCServer.Empty();
        ReleaseItem(NumOPCServer);
        if (OPCServers[NumOPCServer].GroupAdd)
        {
            if (OPCServers[NumOPCServer].SyncRead)  OPCServers[NumOPCServer].OPCSyncIO.Release();
//            if (OPCServers[NumOPCServer].AsyncRead)  OPCServers[NumOPCServer].OPCAsyncIO.Release();
/*            if (OPCServers[NumOPCServer].ConnectPoint)
			{
				OPCServers[NumOPCServer].ipCP->Unadvise(dwAdvise);
				OPCServers[NumOPCServer].ipCP->Release();
			}              */
//            if (OPCServers[NumOPCServer].OnChange) OPCServers[NumOPCServer].ipCallback->Release();
			if (OPCServers[NumOPCServer].GroupAdd)  OPCServers[NumOPCServer].OPCGroup.Release();
			if (OPCServers[NumOPCServer].ServerConnect) OPCServers[NumOPCServer].OPCServer->RemoveGroup(hOPCGroup, FALSE);
            OPCServers[NumOPCServer].hOPCGroup = 0;
        }
        if (OPCServers[NumOPCServer].BrowseServer) OPCServers[NumOPCServer].OPCBrowseServerAdressSpace->Release();
        if (OPCServers[NumOPCServer].ServerConnect) OPCServers[NumOPCServer].OPCServer.Release();
        OPCServers[NumOPCServer].GroupAdd=false;
        OPCServers[NumOPCServer].ServerConnect=false;
//        OPCServers[NumOPCServer].ConnectPoint=false;
//        OPCServers[NumOPCServer].AsyncRead=false;
		OPCServers[NumOPCServer].SyncRead=false;
//        OPCServers[NumOPCServer].OnChange=false;
		OPCServers[NumOPCServer].BrowseServer=false;
    }
    else
    {
		//CLSIDOPCServer.Empty();
        //NameOPCServer.Empty();
        ReleaseItem(-1);
        if (GroupAdd)
        {
            if (SyncRead)  OPCSyncIO.Release();
//            if (AsyncRead)  OPCAsyncIO.Release();
/*            if (ConnectPoint)
			{
				ipCP->Unadvise(dwAdvise);
				ipCP->Release();
			}   */
//			if (OnChange) ipCallback->Release();
			if (GroupAdd)  OPCGroup.Release();
			//if (ServerConnect) OPCServer->RemoveGroup(hOPCGroup, FALSE);
			hOPCGroup = 0;
		}
		if (BrowseServer) OPCBrowseServerAdressSpace->Release();
		if (ServerConnect) OPCServer.Release();
		GroupAdd=false;
		ServerConnect=false;
//		ConnectPoint=false;
//        AsyncRead=false;
		SyncRead=false;
//		OnChange=false;
		BrowseServer=false;
	}
}
//---------------------------------------------------------------------------
void __fastcall TForm1::Button3Click(TObject *Sender)
{
	if (CLSIDOPCServer.m_str)
	{
		Disconnect(SelectedOPCServer);
		StringGrid1->Repaint();
		if (0 == OPCServers[SelectedOPCServer].hOPCGroup)
		SetButtonStatus(false);
		ListItems(SelectedOPCServer);
	}
}
//---------------------------------------------------------------------------

void __fastcall TForm1::FormClose(TObject *Sender, TCloseAction &Action)
{
	LogMessage1=L"Закрытие программы";
	LogMessage2=L"высвобождение ресурсов";
	Log();
	CloseTrendsReadThread=true;
	Sleep(1000);
	for (int i=0;i<NumConnectedOPCServers;i++)
	{
		if (OPCServers[i].ServerConnect) Disconnect(i);
	}
	for (int i = 0; i < NumListTrends; i++)
	{
		TrendsDS[i]->Close();
		TablesDS[i]->Close();
	}
	LogMessage1=L"Закрытие программы";
	LogMessage2=L"ресурсы высвобождены";
	Log();
	LogMessage1=L"Закрытие программы";
	LogMessage2=L"программа закрыта.";
	Log();
}
//---------------------------------------------------------------------------
void __fastcall TForm1::SetButtonStatus(bool Connect)
{
    if (Connect)
    {
        Button2->Enabled=false;
        Button3->Enabled=true;
        StringGrid2->Enabled=true;
        Label4->Caption=AnsiString::AnsiString(OPCServers[SelectedOPCServer].NameOPCServer.m_str);
        Label13->Caption=AnsiString::AnsiString(OPCServers[SelectedOPCServer].NameOPCServer.m_str);
		Label6->Caption=AnsiString::AnsiString(OPCServers[SelectedOPCServer].CLSIDOPCServer.m_str);
	}
    else
    {
        Button2->Enabled=true;
        Button3->Enabled=false;
        Label4->Caption="связь не установлена";
		Label13->Caption="связь не установлена";
        Label6->Caption="не доступно";
        StringGrid2->Enabled=false;
    }
}
//---------------------------------------------------------------------------
void __fastcall TForm1::GetItems(int NumOPCServer)
{
    if (OPCServers[NumOPCServer].ServerConnect)
	{
        char sz2[200];
	    TCHAR szBuffer[256];
    	HRESULT hr = 0;
	    int nTestItem = 0; // how many items there are
    	IEnumString* pEnumString = NULL;
		OPCServers[NumOPCServer].nItems = 0;
    	USES_CONVERSION;
//		StringGrid2->RowCount=100;;
		bool ActualVersion=true;
		int NumUnActual=0;
		if (OPCServers[NumOPCServer].OPCBrowseServerAdressSpace)
		{
			HRESULT hr=OPCServers[NumOPCServer].OPCBrowseServerAdressSpace->BrowseOPCItemIDs(OPC_FLAT, L""/*NULL*/, VT_EMPTY, 0, &pEnumString);
			if (SUCCEEDED(hr))
			{
				LPOLESTR pszName = NULL;
				ULONG count = 0;
				while(((hr = pEnumString->Next(1, &pszName, &count)) == S_OK)&&(OPCServers[NumOPCServer].nItems<MaxParams))
				{
					OPCServers[NumOPCServer].NameOPCItem[OPCServers[NumOPCServer].nItems]=OLE2T(pszName);
					bool ParamOK=false;

					if (NumOPCServer!=-1)
					{

						if (!AdminMode)
						{
							for (int i = 0; i < ListTrends[NumOPCServer].NumVar; i++)
							{
								if (TTrends[NumOPCServer].NameField[i].AnsiCompare(OPCServers[NumOPCServer].NameOPCItem[OPCServers[NumOPCServer].nItems])==0)
								{
									OPCServers[NumOPCServer].AboutOPCItem[OPCServers[NumOPCServer].nItems]=TTrends[NumOPCServer].AboutField[i];
									OPCServers[NumOPCServer].NameDBItem[OPCServers[NumOPCServer].nItems]=TTrends[NumOPCServer].NameFieldDB[i];
									ParamOK=true;
									break;
								}
							}
						}
						else
						{
							ParamOK=true;
						}
					}
					else
					{
						ParamOK=true;
					}
					if (!ParamOK)
					{
						ActualVersion=false;
						NumUnActual++;
					}
					if (ParamOK)
					{
//						StringGrid2->Cells[0][OPCServers[NumOPCServer].nItems]=OPCServers[NumOPCServer].NameOPCItem[OPCServers[NumOPCServer].nItems];
						OPCServers[NumOPCServer].nItems++;
					}
					CoTaskMemFree(pszName);
				}
				pEnumString->Release();
			}
		}
		if (!ActualVersion)
		{
			if (WorkMode!=2)
			{
				String str1="";
				wchar_t s[3];
				_itow(NumUnActual, s, 10);
				str1.operator +=(s);
				str1.operator +=(L" переменных OPC сервера не включены в набор переменных программы");
				Application->MessageBox(str1.c_str()   , L"Предупрждение", MB_OK);
			}
		}
		if (NumOPCServer==0)
		{
			OPCServers[NumOPCServer].AboutOPCItem[OPCServers[NumOPCServer].nItems]=TTrends[NumOPCServer].AboutField[OPCServers[NumOPCServer].nItems];
			OPCServers[NumOPCServer].NameDBItem[OPCServers[NumOPCServer].nItems]=TTrends[NumOPCServer].NameFieldDB[OPCServers[NumOPCServer].nItems];
			OPCServers[NumOPCServer].nItems++;
		}
	}
}
//---------------------------------------------------------------------------

void __fastcall TForm1::Timer1Timer(TObject *Sender)
{
		CurrentTime	= Now();
		DifTime1s=(double)(CurrentTime - LastTime1s);
		if ((!Form1->Time1s1) && (DifTime1s>=s1))
		{
			LastTime1s = Now();
			Form1->Time1s1=true;
		}
		DifTime5s=(double)(CurrentTime - LastTime5s);
		if ((!Form1->Time5s1) && (DifTime5s>=s5))
		{
			LastTime5s = Now();
			Form1->Time5s1=true;
		}
		DifTime60s=(double)(CurrentTime - LastTime60s);
		if ((!Form1->Time60s1) && (DifTime60s>=s60))
		{
			LastTime60s = Now();
			Form1->Time60s1=true;
		}
		DifTimeTrendsReadThread=(double)(CurrentTime - LastTimeTrendsReadThread);
//		double dif=(double)((double)DifTimeTrendsReadThread/(double)s1);
//		Label7->Caption=dif;
		double dif=(double)((double)DifTimeTrendsReadThread/(double)s1);
		DifTimeTrendsReadThread1=dif;
		if (MaxDifTimeTrendsReadThread<dif)
		{
			MaxDifTimeTrendsReadThread=dif;
		}
		Label7->Caption=dif;
		if (dif>60)
		{
			LogMessage1=L"Перезапуск потока TrendsReadThread";
			LogMessage2=L"Время простоя > 60 c";
			Log();
			bool reload=false;
			try
			{
				ThrTrendsRead->Terminate();
			}
			catch(...)
			{
				LogMessage2=L"не может быть уничтожен";
				Log();
			}
			try
			{
				ThrTrendsRead = new TTrendReadThread(true);
			}
			catch(...)
			{
				LogMessage2=L"не может быть выделена память";
				Log();
			}
			try
			{
				ThrTrendsRead = new TTrendReadThread(true);
				ThrTrendsRead->Resume();
				reload=true;
			}
			catch(...)
			{
				LogMessage2=L"не может быть запущен";
				Log();
			}
			if (reload) {
				LogMessage2=L"перезапущен";
				Log();
			}
			else
			{
				LogMessage2=L"не был перезапущен";
				Log();
			}
		}

		if (Time60s1)
		{
/*			if (ThrTrendsRead!=NULL)
			{

/*				if ((ThrTrendsRead->Finished))
				{
					ThrTrendsRead = new TTrendReadThread(true);
					ThrTrendsRead->Resume();
				}*/
/*			}
			else
			{
				ThrTrendsRead = new TTrendReadThread(true);
				ThrTrendsRead->Resume();
			}*/
			Form1->Time60s1=false;
		}
}
//---------------------------------------------------------------------------

void __fastcall TForm1::FormCreate(TObject *Sender)
{
//	return;
	ServerConnect=false;
    GroupAdd=false;
    ConnectPoint=false;
    AsyncRead=false;
    SyncRead=false;
    OnChange=false;
    BrowseServer=false;
    hOPCGroup=0;
    m_dwRate=1000;
    m_fDeadBand=0.0f;
    nItems=0;
    m_dwID=0;
    m_dwCancelID=0;
    TimeRefresh=1000;
    MaxParams=3000;

    BlackColor=clBlack;
    WhiteColor=clWhite;
    GrayColor=clGray;

    Colors[0]=clNavy;
    Colors[1]=clGray;
    Colors[2]=clWhite;//clBlack;
	Colors[3]=clRed;
    Colors[4]=clGreen;
    Colors[5]=clBlue;
    Colors[6]=clPurple;
    Colors[7]=clYellow;
    Colors[8]=clLime;
    Colors[9]=clAqua;
    Colors[10]=clSkyBlue;
    Colors[11]=clMaroon;

    Colors[12]=clNavy;
    Colors[13]=clGray;
    Colors[14]=clWhite;//clBlack;
    Colors[15]=clRed;
    Colors[16]=clGreen;
    Colors[17]=clBlue;
    Colors[18]=clPurple;
    Colors[19]=clYellow;
    Colors[20]=clLime;
    Colors[21]=clAqua;
    Colors[22]=clSkyBlue;
    Colors[23]=clMaroon;

//	FCT=true;
/*	TTimerThread *ThrTimer = new TTimerThread(true);
	TTrendsStoreThread *ThrTrendsStore = new TTrendsStoreThread(true);
	CloseTrendsStoreThread=false;
	CloseTimerThread=false;
	Time1s1=true;
	Time5s1=true;
	Time60s1=true;
	ThrTimer->Resume();
	ThrTrendsStore->Resume();*/
}
//---------------------------------------------------------------------------

void __fastcall TForm1::Button5Click(TObject *Sender)
{
    SelectedOPCServer=-1;
    for (int i=0;i<NumConnectedOPCServers;i++)
    {
		if (OPCServers[i].ServerConnect) Disconnect(i);
    }
	StringGrid2->RowCount=0;
    StringGrid2->Cells[0][0]="";
	StringGrid2->Cells[1][0]="";
	StringGrid2->Cells[2][0]="";
	StringGrid2->Enabled=false;
    Form2->ShowModal();
    StringGrid1->Repaint();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::OnDataChange(
      /* [in] */ DWORD dwTransid,
      /* [in] */ OPCHANDLE hGroup,
      /* [in] */ HRESULT hrMasterquality,
      /* [in] */ HRESULT hrMastererror,
      /* [in] */ DWORD dwCount,
      /* [size_is][in] */ OPCHANDLE __RPC_FAR *phClientItems,
      /* [size_is][in] */ VARIANT __RPC_FAR *pvValues,
      /* [size_is][in] */ WORD __RPC_FAR *pwQualities,
      /* [size_is][in] */ FILETIME __RPC_FAR *pftTimeStamps,
      /* [size_is][in] */ HRESULT __RPC_FAR *pErrors)
{
//    Label14->Caption=Format(_T("V = '%s'"),
//        ARRAYOFCONST((Variant2Str(*pvValues)))).c_str();
}
//---------------------------------------------------------------------------
void __fastcall TForm1::OnReadComplete(
      /* [in] */ DWORD dwTransid,
      /* [in] */ OPCHANDLE hGroup,
      /* [in] */ HRESULT hrMasterquality,
      /* [in] */ HRESULT hrMastererror,
      /* [in] */ DWORD dwCount,
      /* [size_is][in] */ OPCHANDLE __RPC_FAR *phClientItems,
      /* [size_is][in] */ VARIANT __RPC_FAR *pvValues,
      /* [size_is][in] */ WORD __RPC_FAR *pwQualities,
      /* [size_is][in] */ FILETIME __RPC_FAR *pftTimeStamps,
      /* [size_is][in] */ HRESULT __RPC_FAR *pErrors)
{
    Label13->Caption=Format(_T("V = '%s'"),
		ARRAYOFCONST((Variant2Str(*pvValues)))).c_str();
}
//---------------------------------------------------------------------------


void __fastcall TForm1::StringGrid1SelectCell(TObject *Sender, int ACol,
      int ARow, bool &CanSelect)
{
	SelectedRowServer=ARow;
	if (StringGrid1->Cells[0][ARow]!="")
    {
		if (StringGrid1->Cells[0][ARow]!="")
		{
			if (ACol==1)
				ComboBox3->Text = StringGrid1->Cells[0][ARow];
			if (ACol==2)
				ComboBox3->Text = StringGrid1->Cells[1][ARow];
		}
		int NumOPCServer=-1;
		for (int i = 0; i < NumConnectedOPCServers; i++)
		{
			AnsiString ServerCLSID = OPCServers[i].CLSIDOPCServer.m_str;
			if (ServerCLSID.AnsiCompare(StringGrid1->Cells[0][ARow])==0)
			{
				AnsiString ComputerName = OPCServers[i].ComputerName;
				if (ComputerName.AnsiCompare(StringGrid1->Cells[2][ARow])==0)
				{
					NumOPCServer=i;
					break;
				}
			}
		}
		SelectOPCServer(NumOPCServer);
        Button2->Enabled=true;
        Button3->Enabled=false;
		if (OPCServers[NumOPCServer].ServerConnect)
        {
            Button2->Enabled=false;
            Button3->Enabled=true;
        }
    }
	else
    {
        Button2->Enabled=false;
        Button3->Enabled=false;
    }
}
//---------------------------------------------------------------------------
void __fastcall TForm1::ConnectToItems(int NumOPCServer)
{
// Get the item management interface of the group
    CComPtr<IOPCItemMgt> OPCItemMgm;
    OPCITEMDEF ItemDef;
    HRESULT *phResult = NULL;
    HRESULT hr;
    OPCITEMRESULT *OPCItemState = NULL;
    USES_CONVERSION;

    hr=OPCServers[NumOPCServer].OPCGroup->QueryInterface(IID_IOPCItemMgt,
                                            reinterpret_cast<LPVOID*>(&OPCItemMgm));
    if (SUCCEEDED(hr))
    {
//       ItemDef.szItemID = NameOPCItem;
        ItemDef.szAccessPath = T2OLE(_T(""));
        ItemDef.bActive = TRUE;
        ItemDef.hClient = reinterpret_cast<DWORD>(Handle);
        ItemDef.dwBlobSize = 0;
        ItemDef.pBlob = NULL;
		ItemDef.vtRequestedDataType = 0;
		int i=0;
		for (i=0;i<OPCServers[NumOPCServer].nItems;i++)
		{
			ItemDef.szItemID = (wchar_t *)WideString(OPCServers[NumOPCServer].NameOPCItem[i].c_str());
			hr=OPCItemMgm->AddItems(1, &ItemDef, &OPCItemState, &phResult);
			if (SUCCEEDED(hr))
            {
                if (SUCCEEDED(phResult[0]))
                {
                    if (OPCItemState[0].pBlob != NULL)
                    {
                        m_ptrMalloc->Free(OPCItemState[0].pBlob);
                    }

                    m_ptrMalloc->Free(phResult);
					OPCServers[NumOPCServer].hOPCItem[i] = OPCItemState[0].hServer;
                    m_ptrMalloc->Free(OPCItemState);
                }
                else
                {
//                    HRESULT hr = phResult[0];
                    m_ptrMalloc->Free(phResult);
                    m_ptrMalloc->Free(OPCItemState);
                    //OLECHECK((HRESULT)hr);// Always the exception is raised here
                }
            }
        }
    }
}
//---------------------------------------------------------------------------



void __fastcall TForm1::Button1Click(TObject *Sender)
{
    int FindServ=0;
//    if (FindServers(1)>0) FindServ=1;
//    if (FindServers(2)>0) FindServ=1;
//    if (FindServers(3)) FindServ=true;
        FindServ=Form2->FindServers(1);
        if (FindServ>0)
        {
            StringGrid1->Enabled=true;
        }
        else
        {
            if (FindServ<0)
            {
				wchar_t s[2];
				_itow(FindServ, s, 10);
				Application->MessageBox(L"Не удалось соединиться с сервером.", s, MB_OK);
            }
            else
            {
				StringGrid1->RowCount=1;
				StringGrid1->Cells[1][0]="LOPC сервера не найдены";
                Button2->Enabled=false;
            }
        }
    SelectedOPCServer=-1;
    StringGrid1->Repaint();
}
//---------------------------------------------------------------------------





void __fastcall TForm1::FormShow(TObject *Sender)
{
//    Button2->SetFocus();
//Form1->Refresh();

}
//---------------------------------------------------------------------------

void __fastcall TForm1::ComboBox3Change(TObject *Sender)
{
    if (!ServerConnect)
    {
        if (ComboBox3->Text !="")
        {
//            CLSIDOPCServer.Empty();
			CLSIDOPCServer = (wchar_t)(ComboBox3->Text.c_str());
//			CLSIDOPCServer = WideString(ComboBox3->Text.c_str());
			Button2->Enabled=true;
        }
        else
        {
            Button2->Enabled=false;
		}
    }
}
//---------------------------------------------------------------------------

void __fastcall TForm1::CheckBox1Click(TObject *Sender)
{
	TrendsStoreRemoute=CheckBox1->Checked;
}
//---------------------------------------------------------------------------



void __fastcall TForm1::Button4Click(TObject *Sender)
{
	ViewRemouteTrends=true;
	SelectTrendsForm->ShowModal();
}
//---------------------------------------------------------------------------



void __fastcall TForm1::StringGrid1DrawCell(TObject *Sender, int ACol,
	  int ARow, TRect &Rect, TGridDrawState State)
{
	StringGrid1->Canvas->Font->Color=clBlack;
	if (ACol==1)
	{
		if (ARow==SelectedRowServer)//SelectedOPCServer)
			StringGrid1->Canvas->Brush->Color=clGreen;
		else
			StringGrid1->Canvas->Brush->Color=clMedGray;
			StringGrid1->Canvas->FillRect(Rect);
			StringGrid1->Canvas->TextOut(Rect.Left,Rect.Top,StringGrid1->Cells[ACol][ARow]);

	}
	if (ACol==2)
	{
		bool b=false;
		for (int i = 0; i < NumConnectedOPCServers; i++)
		{
			AnsiString ServerCLSID = OPCServers[i].CLSIDOPCServer.m_str;
			if (ServerCLSID.AnsiCompare(StringGrid1->Cells[0][ARow])==0)
			{
				AnsiString ComputerName = OPCServers[i].ComputerName;
				if (ComputerName.AnsiCompare(StringGrid1->Cells[2][ARow])==0)
				{
							if (OPCServers[i].ServerConnect) b=true;
					break;
				}
			}
		}
		if (b)
			StringGrid1->Canvas->Brush->Color=clGreen;
		else
			StringGrid1->Canvas->Brush->Color=clGray;
			StringGrid1->Canvas->FillRect(Rect);
            StringGrid1->Canvas->TextOut(Rect.Left,Rect.Top,StringGrid1->Cells[ACol][ARow]);
	}
}
//---------------------------------------------------------------------------

void __fastcall TForm1::ListItems(int NumOPCServer)
{
	StringGrid2->RowCount=OPCServers[NumOPCServer].nItems;
	StringGrid2->Cells[0][0]="";
	StringGrid2->Cells[1][0]="";
	StringGrid2->Cells[2][0]="";
	StringGrid2->Enabled=true;
	for (int i=0;i<StringGrid2->RowCount;i++)
	{
		StringGrid2->Cells[0][i]=OPCServers[NumOPCServer].AboutOPCItem[i];
		StringGrid2->Cells[1][i]=OPCServers[NumOPCServer].NameOPCItem[i];
		StringGrid2->Cells[2][i]=OPCServers[NumOPCServer].NameOPCItem[i];
    }

}
//---------------------------------------------------------------------------

void __fastcall TForm1::SelectOPCServer(int NumOPCServer)
{
	SelectedOPCServer=NumOPCServer;
	if (NumOPCServer>=0)
	{
		NameOPCServer=OPCServers[NumOPCServer].NameOPCServer.m_str;
		CLSIDOPCServer=OPCServers[NumOPCServer].CLSIDOPCServer.m_str;
		ComboBox1->Text=OPCServers[NumOPCServer].ComputerName;
		if (OPCServers[SelectedOPCServer].ServerConnect)
			SetButtonStatus(true);
		else
			SetButtonStatus(false);
		ListItems(NumOPCServer);
	}
	else
	{
		NameOPCServer=StringGrid1->Cells[1][SelectedRowServer].c_str();
		CLSIDOPCServer=StringGrid1->Cells[0][SelectedRowServer].c_str();
	}
}




void __fastcall TForm1::CheckBox2Click(TObject *Sender)
{
    ShowForm(CheckBox2->Checked);
}
//---------------------------------------------------------------------------

void __fastcall TForm1::ShowForm(bool Advanced)
{
    if (Advanced)
    {
		GroupBox1->Visible=true;
		GroupBox2->Top=GroupBox1->Top+GroupBox1->Height+5;
		GroupBox2->Height=643-GroupBox1->Height-5;//297;
	}
	else
	{
		GroupBox1->Visible=false;
		GroupBox2->Top=24;
		GroupBox2->Height=643;//449;
    }

    Label1->Top=8;
    StringGrid1->Top=24;
    StringGrid1->Height=GroupBox2->Height-136;
    ComboBox3->Top=GroupBox2->Height-105;
    Button2->Top=GroupBox2->Height-73;
    Button3->Top=GroupBox2->Height-73;
    Label3->Top=GroupBox2->Height-41;
    Label4->Top=GroupBox2->Height-41;
    Label5->Top=GroupBox2->Height-21;
    Label6->Top=GroupBox2->Height-21;
}
//---------------------------------------------------------------------------











void __fastcall TForm1::Timer2Timer(TObject *Sender)
{
	LogMessage1=L"Запуск программы";
	LogMessage2=L"запуск основного цикла";
	Log();
		ThrTrendsRead = new TTrendReadThread(true);
		CloseTrendsReadThread=false;
		Time1s1=true;
		Time5s1=true;
		Time60s1=true;
		LastTime1s = Now();
		LastTime5s = Now();
		LastTime60s = Now();
		NumStoredRemoute=0;
		NumStoredLocal=0;
		CurrentTime	= Now();
		int Z=CurrentTime;
		TimeReOpen+=Z;
		LastTimeTrendsReadThread = CurrentTime;
		DifTimeTrendsReadThread=0;
		Timer1->Enabled=true;


		Timer2->Enabled=false;
		ThrTrendsRead->Resume();
	LogMessage1=L"Запуск программы";
	LogMessage2=L"запущен основной цикл";
	Log();

	LogMessage1=L"Запуск программы";
	LogMessage2=L"выполнение загрузочного скрипта";
	Log();
	ExecLoadScript();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::SaveTrendConfig()
{
	int iFileHandle=0;
	int iFileLength;
	int iBytesRead;
    char *pszBuffer;
    AnsiString NameFile=Form1->ListTrends[Form1->SelectedTrend].NameFileConfig;
	AnsiString FullNameFile=ConfigDirectory+NameFile;
	if (FileExists(FullNameFile))
	{
		iFileHandle = FileOpen(FullNameFile,fmOpenWrite);
	}
	else
    {
        iFileHandle = FileCreate(FullNameFile);
//        Str1.operator +=(" не существует");
//        Application->MessageBox(Str1.c_str(), "Сообщение", MB_OK);
	}
	if (iFileHandle)
	{
		FileWrite(iFileHandle, (bool*)&Form1->ListTrends[Form1->SelectedTrend].Visible, sizeof(Form1->ListTrends[Form1->SelectedTrend].Visible));
		FileWrite(iFileHandle, (char*)&Form1->ListTrends[Form1->SelectedTrend].Color, sizeof(Form1->ListTrends[Form1->SelectedTrend].Color));
		int i1=Form1->ListTrends[Form1->SelectedTrend].Y[0];
		int i2=Form1->ListTrends[Form1->SelectedTrend].ScaleY[0];
		int i3=Form1->ListTrends[Form1->SelectedTrend].MinY[0];
		int i4=Form1->ListTrends[Form1->SelectedTrend].MaxY[0];
		FileWrite(iFileHandle, (int*)&Form1->ListTrends[Form1->SelectedTrend].Y, sizeof(Form1->ListTrends[Form1->SelectedTrend].Y));
		FileWrite(iFileHandle, (int*)&Form1->ListTrends[Form1->SelectedTrend].ScaleY, sizeof(Form1->ListTrends[Form1->SelectedTrend].ScaleY ));
		FileWrite(iFileHandle, (int*)&Form1->ListTrends[Form1->SelectedTrend].MinY, sizeof(Form1->ListTrends[Form1->SelectedTrend].MinY));
		FileWrite(iFileHandle, (int*)&Form1->ListTrends[Form1->SelectedTrend].MaxY, sizeof(Form1->ListTrends[Form1->SelectedTrend].MaxY));
		FileClose(iFileHandle);
	}
}
//---------------------------------------------------------------------------

void __fastcall TForm1::LoadTrendConfig()
{
	int c=0;
	if (Form1->SelectedTrend<7)
	{
		for (int j=0;j<3000;j++)
		{
			Form1->ListTrends[Form1->SelectedTrend].Visible[j]=true;
			Form1->ListTrends[Form1->SelectedTrend].Color[j]=c;
			Form1->ListTrends[Form1->SelectedTrend].Y[j]=0;
			Form1->ListTrends[Form1->SelectedTrend].ScaleY[j]=10;
			Form1->ListTrends[Form1->SelectedTrend].MinY[j]=0;
			Form1->ListTrends[Form1->SelectedTrend].MaxY[j]=600;
			c++;
			if (c>=24) c=0;
		}
	}
	if (Form1->SelectedTrend==7)
	{
		for (int j=0;j<3000;j++)
		{
			Form1->ListTrends[Form1->SelectedTrend].Visible[j]=true;
			Form1->ListTrends[Form1->SelectedTrend].Color[j]=c;
			Form1->ListTrends[Form1->SelectedTrend].Y[j]=0;
			Form1->ListTrends[Form1->SelectedTrend].ScaleY[j]=10;
			Form1->ListTrends[Form1->SelectedTrend].MinY[j]=0;
			Form1->ListTrends[Form1->SelectedTrend].MaxY[j]=1000;
			c++;
			if (c>=24) c=0;
		}
	}
	if (Form1->SelectedTrend==8)
	{
		for (int j=0;j<3000;j++)
		{
			Form1->ListTrends[Form1->SelectedTrend].Visible[j]=true;
			Form1->ListTrends[Form1->SelectedTrend].Color[j]=c;
			Form1->ListTrends[Form1->SelectedTrend].Y[j]=0;
			Form1->ListTrends[Form1->SelectedTrend].ScaleY[j]=10;
			Form1->ListTrends[Form1->SelectedTrend].MinY[j]=0;
			Form1->ListTrends[Form1->SelectedTrend].MaxY[j]=15;
			c++;
			if (c>=24) c=0;
		}
	}

	int iFileHandle=0;
	int iFileLength;
	int iBytesRead;
	int iBytesRead1;
	int iBytesRead2;
	char *pszBuffer;
	AnsiString NameFile=Form1->ListTrends[Form1->SelectedTrend].NameFileConfig;
	AnsiString FullNameFile=ConfigDirectory+NameFile;
	if (FileExists(FullNameFile))
	{
		iFileHandle = FileOpen(FullNameFile,fmOpenRead);
	}
	else
	{
//        Str1.operator +=(" не существует");
//        Application->MessageBox(Str1.c_str(), "Сообщение", MB_OK);
	}
	if (iFileHandle)
	{
		iFileLength = FileSeek(iFileHandle,0,2);
		FileSeek(iFileHandle,0,0);
		pszBuffer = new char[iFileLength+1];
		iBytesRead = FileRead(iFileHandle, pszBuffer, iFileLength);
		FileClose(iFileHandle);

		for (int i=0;i<iBytesRead;i++)
		{
			if (i<Form1->ListTrends[Form1->SelectedTrend].MaxParam)
			{
				Form1->ListTrends[Form1->SelectedTrend].Visible[i]=(bool)pszBuffer[i];
				Form1->ListTrends[Form1->SelectedTrend].Visible[i]=(bool)pszBuffer[i];
			}
			if ((i>=Form1->ListTrends[Form1->SelectedTrend].MaxParam)&&(i<Form1->ListTrends[Form1->SelectedTrend].MaxParam*2))
			{
				char c=(char)pszBuffer[i];
				Form1->ListTrends[Form1->SelectedTrend].Color[i-Form1->ListTrends[Form1->SelectedTrend].MaxParam]=c;
			}
			if ((i>=Form1->ListTrends[Form1->SelectedTrend].MaxParam*2)&&(i<Form1->ListTrends[Form1->SelectedTrend].MaxParam*6))
			{
				BYTE c1=(BYTE)pszBuffer[i];
				i++;
				BYTE c2=(BYTE)pszBuffer[i];
				i++;
				BYTE c3=(BYTE)pszBuffer[i];
				i++;
				BYTE c4=(BYTE)pszBuffer[i];
				int i1=(int)(c4*256*256*256+c3*256*256+c2*256+c1);
				int n=(i-Form1->ListTrends[Form1->SelectedTrend].MaxParam*2+1)/4-1;
				Form1->ListTrends[Form1->SelectedTrend].Y[(i-Form1->ListTrends[Form1->SelectedTrend].MaxParam*2+1)/4-1]=i1;
			}
			if ((i>=Form1->ListTrends[Form1->SelectedTrend].MaxParam*6)&&(i<Form1->ListTrends[Form1->SelectedTrend].MaxParam*10))
			{
				BYTE c1=(BYTE)pszBuffer[i];
				i++;
				BYTE c2=(BYTE)pszBuffer[i];
				i++;
				BYTE c3=(BYTE)pszBuffer[i];
				i++;
				BYTE c4=(BYTE)pszBuffer[i];
				int i1=(int)(c4*256*256*256+c3*256*256+c2*256+c1);
				if (i1==0)i1=10;
				int n=(i-Form1->ListTrends[Form1->SelectedTrend].MaxParam*6+1)/4-1;
				Form1->ListTrends[Form1->SelectedTrend].ScaleY[n]=i1;
			}
			if ((i>=Form1->ListTrends[Form1->SelectedTrend].MaxParam*10)&&(i<Form1->ListTrends[Form1->SelectedTrend].MaxParam*14))
			{
				BYTE c1=(BYTE)pszBuffer[i];
				i++;
				BYTE c2=(BYTE)pszBuffer[i];
				i++;
				BYTE c3=(BYTE)pszBuffer[i];
				i++;
				BYTE c4=(BYTE)pszBuffer[i];
				int i1=(int)(c4*256*256*256+c3*256*256+c2*256+c1);
				int n=(i-Form1->ListTrends[Form1->SelectedTrend].MaxParam*10+1)/4-1;
				Form1->ListTrends[Form1->SelectedTrend].MinY[n]=i1;
			}
			if ((i>=Form1->ListTrends[Form1->SelectedTrend].MaxParam*14)&&(i<Form1->ListTrends[Form1->SelectedTrend].MaxParam*18))
			{
				BYTE c1=(BYTE)pszBuffer[i];
				i++;
				BYTE c2=(BYTE)pszBuffer[i];
				i++;
				BYTE c3=(BYTE)pszBuffer[i];
				i++;
				BYTE c4=(BYTE)pszBuffer[i];
				int i1=(int)(c4*256*256*256+c3*256*256+c2*256+c1);
				int n=(i-Form1->ListTrends[Form1->SelectedTrend].MaxParam*14+1)/4-1;
				Form1->ListTrends[Form1->SelectedTrend].MaxY[n]=i1;
			}
		}
		delete [] pszBuffer;
	}
}
//---------------------------------------------------------------------------


void __fastcall TForm1::Label14DblClick(TObject *Sender)
{
	ComboBox3->Enabled=true;
	CheckBox1->Enabled=true;
	CheckBox2->Enabled=true;
}
//---------------------------------------------------------------------------

void __fastcall TForm1::Label15DblClick(TObject *Sender)
{
	ComboBox3->Enabled=false;
	CheckBox1->Enabled=false;
	CheckBox2->Enabled=false;
}
//---------------------------------------------------------------------------

//---------------------------------------------------------------------------

void __fastcall TForm1::LoadServersConfig()
{
	int iFileHandle=0;
	int iFileLength;
	int iBytesRead;
	int iBytesRead1;
	int iBytesRead2;
	char *pszBuffer;
	AnsiString NameFile=Form1->ListTrends[Form1->SelectedTrend].NameFileConfig;
	AnsiString FullNameFile=ConfigDirectory+NameFile;
	if (FileExists(FullNameFile))
	{
		iFileHandle = FileOpen(FullNameFile,fmOpenRead);
	}
	else
	{
//        Str1.operator +=(" не существует");
//        Application->MessageBox(Str1.c_str(), "Сообщение", MB_OK);
    }
    if (iFileHandle)
    {
        iFileLength = FileSeek(iFileHandle,0,2);
        FileSeek(iFileHandle,0,0);
        pszBuffer = new char[iFileLength+1];
		iBytesRead = FileRead(iFileHandle, pszBuffer, iFileLength);
		FileClose(iFileHandle);

		for (int i=0;i<iBytesRead;i++)
		{
			if (i<Form1->ListTrends[Form1->SelectedTrend].MaxParam)
			{
				Form1->ListTrends[Form1->SelectedTrend].Visible[i]=(bool)pszBuffer[i];
				Form1->ListTrends[Form1->SelectedTrend].Visible[i]=(bool)pszBuffer[i];
			}
			if ((i>=Form1->ListTrends[Form1->SelectedTrend].MaxParam)&&(i<Form1->ListTrends[Form1->SelectedTrend].MaxParam*2))
			{
				char c=(char)pszBuffer[i];
				Form1->ListTrends[Form1->SelectedTrend].Color[i-Form1->ListTrends[Form1->SelectedTrend].MaxParam]=c;
			}
			if ((i>=Form1->ListTrends[Form1->SelectedTrend].MaxParam*2)&&(i<Form1->ListTrends[Form1->SelectedTrend].MaxParam*6))
			{
				BYTE c1=(BYTE)pszBuffer[i];
				i++;
				BYTE c2=(BYTE)pszBuffer[i];
				i++;
				BYTE c3=(BYTE)pszBuffer[i];
				i++;
				BYTE c4=(BYTE)pszBuffer[i];
				int i1=(int)(c4*256*256*256+c3*256*256+c2*256+c1);
				int n=(i-Form1->ListTrends[Form1->SelectedTrend].MaxParam*2+1)/4-1;
				Form1->ListTrends[Form1->SelectedTrend].Y[(i-Form1->ListTrends[Form1->SelectedTrend].MaxParam*2+1)/4-1]=i1;
			}
			if ((i>=Form1->ListTrends[Form1->SelectedTrend].MaxParam*6)&&(i<Form1->ListTrends[Form1->SelectedTrend].MaxParam*10))
			{
				BYTE c1=(BYTE)pszBuffer[i];
				i++;
				BYTE c2=(BYTE)pszBuffer[i];
				i++;
				BYTE c3=(BYTE)pszBuffer[i];
				i++;
				BYTE c4=(BYTE)pszBuffer[i];
				int i1=(int)(c4*256*256*256+c3*256*256+c2*256+c1);
				if (i1==0)i1=10;
				int n=(i-Form1->ListTrends[Form1->SelectedTrend].MaxParam*6+1)/4-1;
				Form1->ListTrends[Form1->SelectedTrend].ScaleY[n]=i1;
			}
			if ((i>=Form1->ListTrends[Form1->SelectedTrend].MaxParam*10)&&(i<Form1->ListTrends[Form1->SelectedTrend].MaxParam*14))
			{
				BYTE c1=(BYTE)pszBuffer[i];
				i++;
				BYTE c2=(BYTE)pszBuffer[i];
				i++;
				BYTE c3=(BYTE)pszBuffer[i];
				i++;
				BYTE c4=(BYTE)pszBuffer[i];
				int i1=(int)(c4*256*256*256+c3*256*256+c2*256+c1);
				int n=(i-Form1->ListTrends[Form1->SelectedTrend].MaxParam*10+1)/4-1;
				Form1->ListTrends[Form1->SelectedTrend].MinY[n]=i1;
			}
			if ((i>=Form1->ListTrends[Form1->SelectedTrend].MaxParam*14)&&(i<Form1->ListTrends[Form1->SelectedTrend].MaxParam*18))
			{
				BYTE c1=(BYTE)pszBuffer[i];
				i++;
				BYTE c2=(BYTE)pszBuffer[i];
				i++;
				BYTE c3=(BYTE)pszBuffer[i];
				i++;
				BYTE c4=(BYTE)pszBuffer[i];
				int i1=(int)(c4*256*256*256+c3*256*256+c2*256+c1);
				int n=(i-Form1->ListTrends[Form1->SelectedTrend].MaxParam*14+1)/4-1;
				Form1->ListTrends[Form1->SelectedTrend].MaxY[n]=i1;
			}
		}
		delete [] pszBuffer;
	}
}

void __fastcall TForm1::Label12DblClick(TObject *Sender)
{
	AdminMode ^= true;
}
//---------------------------------------------------------------------------

void __fastcall TForm1::Button6Click(TObject *Sender)
{
	int iFileHandle;

	AnsiString st, st1;
	st="";
	char c1=32;
	char c2=13;
	char c3=10;

	if (!FileExists(".\\grid.txt" ))
	{
		iFileHandle = FileCreate(".\\grid.txt");
	}
	else
	{
		iFileHandle = FileOpen(".\\grid.txt",fmOpenWrite);
	}

	for (int i=1;i<StringGrid2->RowCount;i++)
	{
		st1 = StringGrid2->Cells[1][i];
		st.operator +=(st1);
		st.operator +=(c1);
		st1 = StringGrid2->Cells[2][i];
		st.operator +=(st1);
		st.operator +=(c2);
		st.operator +=(c3);
	}

	FileWrite(iFileHandle, st.c_str(), st.Length());

	FileClose(iFileHandle);
}
//---------------------------------------------------------------------------
void __fastcall TForm1::ReadConfig(AnsiString FileConfig)
{
	WorkMode=0;

	int iFileHandle=0;
	int iFileLength;
	int iBytesRead;
	char *pszBuffer;
	AnsiString
	NameFile=FileConfig;
	AnsiString FullNameFile=ExtractFilePath(Application->ExeName)+NameFile;
	if (FileExists(FullNameFile))
	{
		iFileHandle = FileOpen(FullNameFile,fmOpenRead);
	}
	else
	{
		Application->MessageBox(L"Файл конфигурации программы не найден", L"Сообщение", MB_OK);
    }
	if (iFileHandle)
    {
        iFileLength = FileSeek(iFileHandle,0,2);
		FileSeek(iFileHandle,0,0);
		pszBuffer = new char[iFileLength+1];
		iBytesRead = FileRead(iFileHandle, pszBuffer, iFileLength);
		FileClose(iFileHandle);
	}

	char to;
	to='<';
	char tc;
	tc='>';
	char tcc;
	tcc='/';
	char ch1, ch2;

	int MaxLevel=10;
	int Level=0;

	AnsiString StInfo="";
	int LenInfo;
	AnsiString StTag="";
	int LenTag;

	bool SostReadTagOpen=false;
	bool SostReadTagClose=false;
	bool SostReadInfo=false;

	AnsiString VersionCfg="";
	AnsiString AboutCfg="";

	int WorkModeConfig=0;
	bool ServerModeConfig=false;

	int CurCfg=0;
	int NumCfg=0;

	for (int c=0; c < iBytesRead; c++)
	{
		ch1=pszBuffer[c];
		if ((c+1)<iBytesRead) ch2=pszBuffer[c+1];
		for (int i=0;i<Level;i++)
		{
//				StInfo.operator +=(ch1);
//				LenInfo++;
		}
		if (SostReadTagOpen||SostReadTagClose)
		{
			if ((ch1!=to)&&(ch1!=tc)&&(ch1!=tcc))
			{
				StTag.operator +=(ch1);
				LenTag++;
			}
		}
		if (SostReadInfo)
		{
			if ((ch1!=to)&&(ch1!=tc)&&(ch1!=tcc))
			{
				StInfo.operator +=(ch1);
				LenInfo++;
			}
		}
		if (ch1==to)
		{
			if (ch2==tcc)
			{
				c++;
				SostReadTagClose=true;
				SostReadInfo=false;
//				String St1=StInfo.c_str();
//				String St2=StTag.c_str();
//				Application->MessageBox(St1.c_str(), St2.c_str(), MB_OK);
				LenTag=0;
				StTag="";
			}
			else
			{
				if (SostReadInfo) Level++;
				if (Level>=MaxLevel) Level=MaxLevel-1;
				SostReadTagOpen=true;
				SostReadInfo=false;
				LenTag=0;
				StTag="";
			}
		}
		if (ch1==tc)
		{
			if (SostReadTagClose)
			{
				SostReadTagClose=false;
				String St=StTag.c_str();
				if (!StTag.AnsiCompare("ver"))
				{
					VersionCfg=StInfo;
				}
				if (!StTag.AnsiCompare("about"))
				{
					AboutCfg=StInfo;
				}
				if (NumCfg>0)
				{
					if (!StTag.AnsiCompare("svrmode"))
					{
						if (!StInfo.AnsiCompare("Сервер"))
						{
							ServerModeConfig=true;
						}
						else
						{
							ServerModeConfig=false;
						}
					}
					if (!StTag.AnsiCompare("workmode"))
					{
						WorkModeConfig=0;
						if (!StInfo.AnsiCompare("Котлы"))
						{
							WorkModeConfig=1;
						}
						if (!StInfo.AnsiCompare("Турбины"))
						{
							WorkModeConfig=2;
						}
					}
					Level--;
					if (Level<0) Level=0;
				}
			}
			if (SostReadTagOpen)
			{
				SostReadTagOpen=false;
				SostReadInfo=true;
				LenInfo=0;
				StInfo="";
				if (!StTag.AnsiCompare("cfg"))
				{
					NumCfg++;
					CurCfg=NumCfg-1;
					WorkModeConfig=0;
					ServerModeConfig=false;
				}
			}
		}
	}
	LogMessage1=L"Версия конфигурации";
	LogMessage2=VersionCfg;
	Log();
	LogMessage1=L"Конфигурация";
	LogMessage2=AboutCfg;
	Log();

	WorkMode=WorkModeConfig;
	ServerMode=ServerModeConfig;
}
//---------------------------------------------------------------------------

void __fastcall TForm1::Button7Click(TObject *Sender)
{
/*	if (!ServerStarted)
	{

		InitWTOPCsvr ((BYTE *)&CLSID_SBSOPCServer,1000);

		::UpdateRegistry(
						(BYTE *)&CLSID_SBSOPCServer,
						SBSOPCServerName.c_str(),
						SBSOPCServerAbout.c_str(),
						SBSHelpPath.c_str());
		ServerStarted=true;

		NumSBSOPCParam=0;
		for (int i=0;i<NumValidOPCServers;i++)
		{
			int id=ValidOPCServers[i].idTrend;
			for (int j=0;j<ListTrends[id].NumVar;j++)
			{
				SBSOPCParam[NumSBSOPCParam] = new TSBSOPCParam;

				AnsiString st="";
				wchar_t s[4];
				_itow (i, s, 10);
				st.operator +=(s);
				st.operator +=(".");
				st.operator +=(TTrends[id].NameField[j]);

				for (int k=0; k<4; k++)
				{
					SBSOPCParam[NumSBSOPCParam]->alarms[k] = 0.0;
					SBSOPCParam[NumSBSOPCParam]->severity[k] = 200;
					SBSOPCParam[NumSBSOPCParam]->enabled[k] = FALSE;
				}
				SBSOPCParam[NumSBSOPCParam]->alarms[2] = 100.0;
				SBSOPCParam[NumSBSOPCParam]->alarms[3] = 100.0;
				SBSOPCParam[NumSBSOPCParam]->Value.vt = VT_R4;
				SBSOPCParam[NumSBSOPCParam]->Value.fltVal = 0.0;

				SBSOPCParam[NumSBSOPCParam]->Name = st;
				SBSOPCParam[NumSBSOPCParam]->Handle = CreateTag(
													(LPCSTR)SBSOPCParam[NumSBSOPCParam]->Name.c_str(),
													SBSOPCParam[NumSBSOPCParam]->Value,
													OPC_QUALITY_GOOD,
													true);

				SetItemLevelAlarm (SBSOPCParam[NumSBSOPCParam]->Handle, ID_LOLO_LIMIT, SBSOPCParam[NumSBSOPCParam]->alarms[0], SBSOPCParam[NumSBSOPCParam]->severity[0], SBSOPCParam[NumSBSOPCParam]->enabled[0]);
				SetItemLevelAlarm (SBSOPCParam[NumSBSOPCParam]->Handle, ID_LO_LIMIT, SBSOPCParam[NumSBSOPCParam]->alarms[1], SBSOPCParam[NumSBSOPCParam]->severity[1], SBSOPCParam[NumSBSOPCParam]->enabled[1]);
				SetItemLevelAlarm (SBSOPCParam[NumSBSOPCParam]->Handle, ID_HI_LIMIT, SBSOPCParam[NumSBSOPCParam]->alarms[2], SBSOPCParam[NumSBSOPCParam]->severity[2], SBSOPCParam[NumSBSOPCParam]->enabled[2]);
				SetItemLevelAlarm (SBSOPCParam[NumSBSOPCParam]->Handle, ID_HIHI_LIMIT, SBSOPCParam[NumSBSOPCParam]->alarms[3], SBSOPCParam[NumSBSOPCParam]->severity[3], SBSOPCParam[NumSBSOPCParam]->enabled[3]);

				NumSBSOPCParam++;
			}
		}



	}         */
}
//---------------------------------------------------------------------------

void __fastcall TForm1::CheckBox3Click(TObject *Sender)
{
	ShowValues=CheckBox3->Checked;
}
//---------------------------------------------------------------------------

void __fastcall TForm1::Button8Click(TObject *Sender)
{
//	LPTSTR User;
//	LPTSTR Computer;
//	char User[255];
//	char Computer[255];
//	DWORD l=0;
//	GetUserName(User, &l);
//	GetComputerName(Computer, &l);
//	GetUserName(User, &l);

//		TestThread = new TTestThread(true);
//		TestThread->Resume();

}
//---------------------------------------------------------------------------

void __fastcall TForm1::Button9Click(TObject *Sender)
{
	CloseTestThread=true;
	CloseTrendsReadThread=true;
}
//---------------------------------------------------------------------------

void __fastcall TForm1::Log()
{
	int iFileHandle;

	AnsiString st;
	st="";
	char c0=9;
	char c1=32;
	char c2=13;
	char c3=10;
	AnsiString path=ProgramDirectory+"\\log.txt";
//	if (!FileExists(".\\log.txt" ))
	if (!FileExists(path))
	{
//		iFileHandle = FileCreate(".\\log.txt");
		iFileHandle = FileCreate(path);
	}
	else
	{
//		iFileHandle = FileOpen(".\\log.txt",fmOpenWrite);
		iFileHandle = FileOpen(path,fmOpenWrite);
	}
//	TDateTime CurrentTime;
//		st1 = CurrentTime.DateTimeString();
		st.operator +=(Now().DateTimeString());
		st.operator +=(c0);
		st.operator +=(UserMessage);
		st.operator +=(c0);
		st.operator +=(LogMessage1);
		st.operator +=(c0);
		st.operator +=(LogMessage2);
		st.operator +=(c2);
		st.operator +=(c3);
//	}
	FileSeek(iFileHandle, 0, 2);
	FileWrite(iFileHandle, st.c_str(), st.Length());

	FileClose(iFileHandle);
}


void __fastcall TForm1::Button10Click(TObject *Sender)
{
/*	if (ServerStarted)
	{
		ServerStarted=false;
		UninitWTOPCsvr ();
		UnregisterServer ((BYTE *)&CLSID_SBSOPCServer, SBSOPCServerName.c_str());
		for (int i=0; i < NumSBSOPCParam; i++)
		{
			delete (SBSOPCParam[i]);
		}
	}*/
}
//---------------------------------------------------------------------------

void __fastcall TForm1::Button11Click(TObject *Sender)
{
	MnemoForm->Init("");
	MnemoForm->ShowModal();
}
//---------------------------------------------------------------------------

//---------------------------------------------------------------------------
void __fastcall TForm1::ExecLoadScript()
{
	int iFileHandle=0;
	int iFileLength;
	int iBytesRead;
	char *pszBuffer;
	AnsiString NameFile;
	NameFile="load.scr";
	if (WorkMode==1)
	{
		if (ServerMode)
		{
			NameFile="load_k_s.scr";
			CheckBox1->Enabled=true;
			CheckBox4->Enabled=true;
		}
		else
		{
			NameFile="load_k_c.scr";
			CheckBox1->Enabled=false;
			CheckBox4->Enabled=false;
		}
	}
	if (WorkMode==2)
	{
		if (ServerMode)
		{
			NameFile="load_t_s.scr";
			CheckBox1->Enabled=true;
			CheckBox4->Enabled=true;
		}
		else
		{
			NameFile="load_t_c.scr";
			CheckBox1->Enabled=false;
			CheckBox4->Enabled=false;
		}
	}
	AnsiString FullNameFile=ExtractFilePath(Application->ExeName)+NameFile;
	if (FileExists(FullNameFile))
	{
		iFileHandle = FileOpen(FullNameFile,fmOpenRead);
	}
	else
	{
		Application->MessageBox(L"Файл загрузочного скрипта не найден", L"Сообщение", MB_OK);
    }
	if (iFileHandle)
	{
		iFileLength = FileSeek(iFileHandle,0,2);
		FileSeek(iFileHandle,0,0);
		pszBuffer = new char[iFileLength+1];
		iBytesRead = FileRead(iFileHandle, pszBuffer, iFileLength);
		FileClose(iFileHandle);
	}

	char to;
	to='<';
	char tc;
	tc='>';
	char tcc;
	tcc='/';
	char ch1, ch2;

	int MaxLevel=10;
	int Level=0;

	AnsiString StInfo="";
	int LenInfo;
	AnsiString StTag="";
	int LenTag;

	bool SostReadTagOpen=false;
	bool SostReadTagClose=false;
	bool SostReadInfo=false;

	AnsiString ExecuteTypeScript=L"Выполнение скрипта";
	AnsiString ConnectTypeScript=L"Подключиться к серверу";
	AnsiString ShowTypeScript=L"Вывести мнемосхему";
	AnsiString ShowValuesTypeScript=L"Показ значений";
	AnsiString StoreTrendsTypeScript=L"Ведение трендов";

	AnsiString VersionScript="";
	AnsiString AboutScript="";

	int CurScr=0;
	int NumScr=0;

	for (int c=0; c < iBytesRead; c++)
	{
		ch1=pszBuffer[c];
		if ((c+1)<iBytesRead) ch2=pszBuffer[c+1];
		for (int i=0;i<Level;i++)
		{
//				StInfo.operator +=(ch1);
//				LenInfo++;
		}
		if (SostReadTagOpen||SostReadTagClose)
		{
			if ((ch1!=to)&&(ch1!=tc)&&(ch1!=tcc))
			{
				StTag.operator +=(ch1);
				LenTag++;
			}
		}
		if (SostReadInfo)
		{
			if ((ch1!=to)&&(ch1!=tc)&&(ch1!=tcc))
			{
				StInfo.operator +=(ch1);
				LenInfo++;
			}
		}
		if (ch1==to)
		{
			if (ch2==tcc)
			{
				c++;
				SostReadTagClose=true;
				SostReadInfo=false;
//				String St1=StInfo.c_str();
//				String St2=StTag.c_str();
//				Application->MessageBox(St1.c_str(), St2.c_str(), MB_OK);
				LenTag=0;
				StTag="";
			}
			else
			{
				if (SostReadInfo) Level++;
				if (Level>=MaxLevel) Level=MaxLevel-1;
				SostReadTagOpen=true;
				SostReadInfo=false;
				LenTag=0;
				StTag="";
			}
		}
		if (ch1==tc)
		{
			if (SostReadTagClose)
			{
				SostReadTagClose=false;
				String St=StTag.c_str();
				if (!StTag.AnsiCompare("ver"))
				{
					VersionScript=StInfo;
				}
				if (!StTag.AnsiCompare("about"))
				{
					AboutScript=StInfo;
				}
//				if (NumMnemo>0)
//				{
					if (!StTag.AnsiCompare("cmdscr"))
					{
						Scripts[CurScr].CommandScript=StInfo;
					}
					if (!Scripts[CurScr].CommandScript.AnsiCompare(ConnectTypeScript))
					{
						if (!StTag.AnsiCompare("nsvr"))
						{
							Scripts[CurScr].NServerConnect=StInfo.ToInt();
						}
					}
					if (!Scripts[CurScr].CommandScript.AnsiCompare(ShowValuesTypeScript))
					{
						if (!StTag.AnsiCompare("showval"))
						{
							Scripts[CurScr].ShowValues=StInfo;
						}
					}
					if (!Scripts[CurScr].CommandScript.AnsiCompare(StoreTrendsTypeScript))
					{
						if (!StTag.AnsiCompare("storeval"))
						{
							Scripts[CurScr].StoreTrends=StInfo;
						}
					}
					if (!Scripts[CurScr].CommandScript.AnsiCompare(ExecuteTypeScript))
					{
						if (!StTag.AnsiCompare("execmdscr"))
						{
							Scripts[CurScr].ExecuteScript=StInfo;
						}
					}
					if (!Scripts[CurScr].CommandScript.AnsiCompare(ShowTypeScript))
					{
						if (!StTag.AnsiCompare("namemnemo"))
						{
							Scripts[CurScr].NameMnemoShow=StInfo;
						}
					}
//				}
				Level--;
				if (Level<0) Level=0;
			}
			if (SostReadTagOpen)
			{
				SostReadTagOpen=false;
				SostReadInfo=true;
				LenInfo=0;
				StInfo="";
//				String St=StTag.c_str();
//				Application->MessageBox(St.c_str(), L"Открывается тег", MB_OK);
				if (!StTag.AnsiCompare("scr"))
				{
					NumScr++;
					CurScr=NumScr-1;

					Scripts[CurScr].CommandScript="";
					Scripts[CurScr].ExecuteScript="";
					Scripts[CurScr].NServerConnect=0;
					Scripts[CurScr].ShowValues="";
					Scripts[CurScr].StoreTrends="";
					Scripts[CurScr].NameMnemoShow="";
				}
			}
		}
	}
	LogMessage1=L"Версия скрипта";
	LogMessage2=VersionScript;
	Log();
	LogMessage1=L"Скрипт";
	LogMessage2=AboutScript;
	Log();
	for (int i=0; i < NumScr; i++)
	{
		if (!Scripts[i].CommandScript.AnsiCompare(ConnectTypeScript))
		{
			LogMessage1=Scripts[i].CommandScript;
			LogMessage2=OPCServers[Scripts[i].NServerConnect].ComputerName+
						" : "+
						AnsiString::AnsiString(OPCServers[Scripts[i].NServerConnect].NameOPCServer)+
						" : "+
						AnsiString::AnsiString(OPCServers[Scripts[i].NServerConnect].CLSIDOPCServer);
			Log();
			try
			{
				SelectOPCServer(Scripts[i].NServerConnect);
				int ConnectResult=0;
				ConnectResult=ConnectToServer(Scripts[i].NServerConnect);//(SelectedOPCServer);
				StringGrid1->Repaint();
				if (ConnectResult<0)
				{
					wchar_t s[2];
					_itow (ConnectResult, s, 10);
					LogMessage1=Scripts[i].CommandScript;
					LogMessage2=L"Не удалось соединиться с сервером.";
					SetButtonStatus(false);
				}
				else
				{
					SetButtonStatus(true);
				}
			}
			catch(...)
			{
				Application->MessageBox(L"Error", L"Error", MB_OK);
				Disconnect(-1);
				SetButtonStatus(false);
				throw;
			}
		}
		if (!Scripts[i].CommandScript.AnsiCompare(ExecuteTypeScript))
		{
			LogMessage1=Scripts[i].CommandScript;
			LogMessage2=Scripts[i].ExecuteScript;
			Log();
			if (Scripts[i].ExecuteScript.AnsiCompare("Выполнять скрипт")) break;
		}
		if (!Scripts[i].CommandScript.AnsiCompare(ShowTypeScript))
		{
			LogMessage1=Scripts[i].CommandScript;
			LogMessage2=Scripts[i].NameMnemoShow;
			Log();
			MnemoForm->Init(Scripts[i].NameMnemoShow);
			MnemoForm->ShowModal();
		}
		if (!Scripts[i].CommandScript.AnsiCompare(ShowValuesTypeScript))
		{
			LogMessage1=Scripts[i].CommandScript;
			LogMessage2=Scripts[i].ShowValues;
			Log();
			if (!Scripts[i].ShowValues.AnsiCompare("Показывать значения"))
			{
				ShowValues=true;
			}
			else
			{
				ShowValues=false;
			}
			CheckBox3->Checked=ShowValues;
		}
		if (!Scripts[i].CommandScript.AnsiCompare(StoreTrendsTypeScript))
		{
			LogMessage1=Scripts[i].CommandScript;
			LogMessage2=Scripts[i].ShowValues;
			Log();
			if (!Scripts[i].StoreTrends.AnsiCompare("Вести тренды"))
			{
				TrendsStoreRemoute=true;
			}
			else
			{
				TrendsStoreRemoute=false;
			}
			CheckBox1->Checked=TrendsStoreRemoute;
		}
	}
}
//---------------------------------------------------------------------------

//---------------------------------------------------------------------------
void __fastcall TForm1::LoadServersList()
{
	int iFileHandle=0;
	int iFileLength;
	int iBytesRead;
	char *pszBuffer;
	AnsiString
	NameFile="servers.lst";
	if (WorkMode==1) NameFile="servers_k.lst";
	if (WorkMode==2) NameFile="servers_t.lst";
	AnsiString FullNameFile=ExtractFilePath(Application->ExeName)+NameFile;
	if (FileExists(FullNameFile))
	{
		iFileHandle = FileOpen(FullNameFile,fmOpenRead);
	}
	else
	{
		Application->MessageBox(L"Файл списка серверов не найден", L"Сообщение", MB_OK);
    }
	if (iFileHandle)
    {
        iFileLength = FileSeek(iFileHandle,0,2);
		FileSeek(iFileHandle,0,0);
		pszBuffer = new char[iFileLength+1];
		iBytesRead = FileRead(iFileHandle, pszBuffer, iFileLength);
		FileClose(iFileHandle);
	}

	char to;
	to='<';
	char tc;
	tc='>';
	char tcc;
	tcc='/';
	char ch1, ch2;

	int MaxLevel=10;
	int Level=0;

	AnsiString StInfo="";
	int LenInfo;
	AnsiString StTag="";
	int LenTag;

	bool SostReadTagOpen=false;
	bool SostReadTagClose=false;
	bool SostReadInfo=false;

	AnsiString VersionList="";
	AnsiString AboutList="";

	int CurSvr=0;
	int NumSvr=0;

	for (int c=0; c < iBytesRead; c++)
	{
		ch1=pszBuffer[c];
		if ((c+1)<iBytesRead) ch2=pszBuffer[c+1];
		for (int i=0;i<Level;i++)
		{
//				StInfo.operator +=(ch1);
//				LenInfo++;
		}
		if (SostReadTagOpen||SostReadTagClose)
		{
			if ((ch1!=to)&&(ch1!=tc)&&(ch1!=tcc))
			{
				StTag.operator +=(ch1);
				LenTag++;
			}
		}
		if (SostReadInfo)
		{
			if ((ch1!=to)&&(ch1!=tc)&&(ch1!=tcc))
			{
				StInfo.operator +=(ch1);
				LenInfo++;
			}
		}
		if (ch1==to)
		{
			if (ch2==tcc)
			{
				c++;
				SostReadTagClose=true;
				SostReadInfo=false;
//				String St1=StInfo.c_str();
//				String St2=StTag.c_str();
//				Application->MessageBox(St1.c_str(), St2.c_str(), MB_OK);
				LenTag=0;
				StTag="";
			}
			else
			{
				if (SostReadInfo) Level++;
				if (Level>=MaxLevel) Level=MaxLevel-1;
				SostReadTagOpen=true;
				SostReadInfo=false;
				LenTag=0;
				StTag="";
			}
		}
		if (ch1==tc)
		{
			if (SostReadTagClose)
			{
				SostReadTagClose=false;
				String St=StTag.c_str();
				if (!StTag.AnsiCompare("ver"))
				{
					VersionList=StInfo;
				}
				if (!StTag.AnsiCompare("about"))
				{
					AboutList=StInfo;
				}
				if (NumSvr>0)
				{
						if (!StTag.AnsiCompare("namesvr"))
						{
							LoadedServers[CurSvr].NameServer=StInfo;
						}
						if (!StTag.AnsiCompare("aboutsvr"))
						{
							LoadedServers[CurSvr].AboutServer=StInfo;
						}
						if (!StTag.AnsiCompare("clsid"))
						{
							LoadedServers[CurSvr].CLSIDServer=StInfo;
						}
						if (!StTag.AnsiCompare("compname"))
						{
							LoadedServers[CurSvr].ComputerNameServer=StInfo;
						}
						if (!StTag.AnsiCompare("nametrend"))
						{
							LoadedServers[CurSvr].NameTrendServer=StInfo;
						}
						if (!StTag.AnsiCompare("conftrend"))
						{
							LoadedServers[CurSvr].ConfigTrendServer=StInfo;
						}
						if (!StTag.AnsiCompare("pathtrend"))
						{
							LoadedServers[CurSvr].PathTrendServer=StInfo;
						}
						if (!StTag.AnsiCompare("filetrend"))
						{
							LoadedServers[CurSvr].FileTrendServer=StInfo;
						}
						if (!StTag.AnsiCompare("numvar"))
						{
							LoadedServers[CurSvr].NumVarServer=StInfo.ToInt();
						}
					Level--;
					if (Level<0) Level=0;
				}
			}
			if (SostReadTagOpen)
			{
				SostReadTagOpen=false;
				SostReadInfo=true;
				LenInfo=0;
				StInfo="";
//				String St=StTag.c_str();
//				Application->MessageBox(St.c_str(), L"Открывается тег", MB_OK);
				if (!StTag.AnsiCompare("svr"))
				{
					NumSvr++;
					CurSvr=NumSvr-1;
					LoadedServers[CurSvr].NameServer="";
					LoadedServers[CurSvr].CLSIDServer="";
					LoadedServers[CurSvr].AboutServer="";
					LoadedServers[CurSvr].ComputerNameServer="";
					LoadedServers[CurSvr].NameTrendServer="";
					LoadedServers[CurSvr].ConfigTrendServer="";
					LoadedServers[CurSvr].PathTrendServer="";
					LoadedServers[CurSvr].FileTrendServer="";
					LoadedServers[CurSvr].NumVarServer=0;
				}
			}
		}
	}
	LogMessage1=L"Версия списка";
	LogMessage2=VersionList;
	Log();
	LogMessage1=L"Список";
	LogMessage2=AboutList;
	Log();

	for (int i=0; i < NumSvr; i++)
	{
		TrendPath[i]=LoadedServers[i].PathTrendServer;
		TrendFile[i]=LoadedServers[i].FileTrendServer;
		TrendFilePath[i]=TrendPath[i];
		TrendFilePath[i].operator +=(TrendFile[i]);
		TrendFilePath[i].operator +=("_trends.now");

		NumListTrends=i+1;
		ListTrends[i].TrendName=LoadedServers[i].NameTrendServer;
		ListTrends[i].NameFileConfig=LoadedServers[i].ConfigTrendServer;
		ListTrends[i].NumVar=LoadedServers[i].NumVarServer;
		ListTrends[i].NTrend=i;
		ListTrends[i].MaxParam=3000;
		ListTrends[i].TableName="table";
		ListTrends[i].TableNamePrefix[0]="_0";
		ListTrends[i].TableNamePrefix[1]="_1";
		ListTrends[i].TableNamePrefix[2]="_2";
		ListTrends[i].TableNamePrefix[3]="_3";
		ListTrends[i].TableNamePrefix[4]="_4";
		ListTrends[i].TableNamePrefix[5]="_5";
		ListTrends[i].TableNamePrefix[6]="_6";
		ListTrends[i].TableNamePrefix[7]="_7";
		ListTrends[i].TableNamePrefix[8]="_8";
		ListTrends[i].TableNamePrefix[9]="_9";

		int c=0;
		for (int j=0;j<ListTrends[i].MaxParam;j++)
		{
			ListTrends[i].Visible[j]=true;
			ListTrends[i].Color[j]=c;
			c++;
			if (c>=24) c=0;
		}

		NumValidOPCServers=i+1;
		ValidOPCServers[i].NameField=LoadedServers[i].NameServer;
		ValidOPCServers[i].AboutField=LoadedServers[i].AboutServer;
		ValidOPCServers[i].CLSID=LoadedServers[i].CLSIDServer;
		ValidOPCServers[i].ComputerName=LoadedServers[i].ComputerNameServer;
		ValidOPCServers[i].idTrend=i;


		NumConnectedOPCServers=i+1;
		OPCServers[i].NameOPCServer=ValidOPCServers[i].NameField.c_str();
		OPCServers[i].CLSIDOPCServer=ValidOPCServers[i].CLSID.c_str();
		OPCServers[i].ComputerName=ValidOPCServers[i].ComputerName;
		OPCServers[i].idTrend=ValidOPCServers[i].idTrend;
		OPCServers[i].ServerConnect=false;
	}
}
//---------------------------------------------------------------------------


void __fastcall TForm1::CheckBox4Click(TObject *Sender)
{
	TrendsStoreLocal=CheckBox4->Checked;
}
//---------------------------------------------------------------------------


void __fastcall TForm1::FormResize(TObject *Sender)
{
Form1->Refresh();

}
//---------------------------------------------------------------------------

void __fastcall TForm1::Button12Click(TObject *Sender)
{
	ViewRemouteTrends=false;
	SelectTrendsForm->ShowModal();
}
//---------------------------------------------------------------------------

