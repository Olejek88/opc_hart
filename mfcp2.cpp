#define _CRT_SECURE_NO_WARNINGS
#define _WIN32_DCOM				// Enables DCOM extensions
#define INITGUID				// Initialize OLE constants

//#include <atlbase.h>
#include <stdio.h>
#include <math.h>				// some mathematical function
//#include "afxcmn.h"				// MFC function such as CString,....
#include "hartbus.h"			// no comments
#include "hartdrv.h"			// no comments
#include "unilog.h"				// universal utilites for creating log-files
#include <locale.h>				// set russian codepage
#include "opcda.h"				// basic function for OPC:DA
#include "lightopc.h"			// light OPC library header file
#include "serialport.h"			// function for work with Serial Port

#define ECL_SID "opc.hart"		// identificator of OPC server
#define TAGNAME_LEN 64			// tag lenght

// OPC vendor info (Major/Minor/Build/Reserv)
static const loVendorInfo vendor = {0,92,1,0,"Hart Universal OPC Server" };
loService *my_service;			// name of light OPC Service
static void a_server_finished(void*, loService*, loClient*);// OnServer finish his work
static int OPCstatus=OPC_STATUS_RUNNING;					// status of OPC server
//---------------------------------------------------------------------------------
double cIEEE754toFloat (unsigned char *Data);	// convert from IEEE754(as unsigned char) to float 
double cIEEE754toFloatC (char *Data);			// convert from IEEE754(as char) to float 
char* cFloatCtoIEEE754 (float num);				// convert from float to char[4] IEEE754
char* Base64toString (char *Data,int lenght);	// convert from base64 to string
char* StringtoBase64 (char *Data,int lenght);	// convert from string to base64
char* cDharttoDate (char *Data);				// convert from hart Date format to DATE
char* ReadParam (char *SectionName,char *Value);// read parametr from .ini file
long c24bitto32bit (char *Data);				// convert from 24-bit to long type (32-bit)
int tTotal=100;					// total quantity of tags
int comp=0,speed=9600;			// com-port settings
static Hart *devp;				// pointer to tag structure
static loTagId ti[100];			// Tag id
static loTagValue tv[100];		// Tag value
static char *tn[100];			// Tag name
char argv0[FILENAME_MAX + 32];	// lenght of command line (file+path (260+32))
unilog *logg=NULL;				// new structure of unilog
static OPCcfg cfg;				// new structure of cfg
SerialPort port;				// com-port
void addCommToPoll();			// add commands to read list
UINT ScanBus();					// bus scanned programm
UINT PollDevice(int device);	// polling single device
UINT InitDriver();				// func of initialising port and creating service
UINT DestroyDriver();			// function of detroying driver and service
HRESULT WriteDevice(int device,const unsigned cmdnum,LPSTR data);	// write tag to device
FILE *CfgFile;						// pointer to .ini file
static CRITICAL_SECTION lk_values;	// protects ti[] from simultaneous access 
CRITICAL_SECTION drv_access;
BOOL	DTRHIGH = FALSE;		// DTR on Write Low or High

static int mymain(HINSTANCE hInstance, int argc, char *argv[]);
static int show_error(LPCTSTR msg);		// just show messagebox with error
static int show_msg(LPCTSTR msg);		// just show messagebox with message
static void poll_device(void);			// function polling device
void dataToTag(int device);				// copy data buffer to tag

// 5b9d02df-0e29-407a-8571-c3c315bc8b50 
DEFINE_GUID(UHSID_hOPCserverDll,		// create 128-bit code for sys
    0x5b9d02df,
    0x0e29,
    0x407a,
    0x85, 0x71, 0xc3, 0xc3, 0x15, 0xbc, 0x8b, 0x50
  );
DEFINE_GUID(UHSID_hOPCserverExe,
    0x59cabf40,
    0x0c72,
    0x48c1,
    0x9e, 0x2c, 0x4c, 0x69, 0xca, 0xf9, 0x19, 0x63
  );
//---------------------------------------------------------------------------------
void timetToFileTime( time_t t, LPFILETIME pft )
{
    LONGLONG ll = Int32x32To64(t, 10000000) + 116444736000000000;
    pft->dwLowDateTime = (DWORD) ll;
    pft->dwHighDateTime =  (unsigned long)(ll >>32);	//(unsigned long)
}

char *absPath(char *fileName)					// return abs path of file
{
  static char path[sizeof(argv0)]="\0";			// path - massive of comline
  char *cp;										
  
  if(*path=='\0') strcpy(path, argv0);			
  if(NULL==(cp=strrchr(path,'\\'))) cp=path; else cp++;
  cp=strcpy(cp,fileName);
  return path;
}

inline void init_common(void)		// create log-file
{
  logg = unilog_Create(ECL_SID, absPath(LOG_FNAME), NULL, 0, ll_DEBUG); // level [ll_FATAL...ll_DEBUG] 
  UL_INFO((LOGID, "Start"));		// write in log
}

inline void cleanup_common(void)	// delete log-file
{
  UL_INFO((LOGID, "Finish"));
  unilog_Delete(logg); logg = NULL; // + logs was not currently
}

int show_error(LPCTSTR msg)			// just show messagebox with error
{
  ::MessageBox(NULL, msg, ECL_SID, MB_ICONSTOP|MB_OK);
  return 1;
}
int show_msg(LPCTSTR msg)			// just show messagebox with message
{
  ::MessageBox(NULL, msg, ECL_SID, MB_OK);
  return 1;
}
inline void cleanup_all(DWORD objid)
{
// Informs OLE that a class object, previously registered is no longer available for use
  if (FAILED(CoRevokeClassObject(objid)))  UL_WARNING((LOGID, "CoRevokeClassObject() failed..."));
  DestroyDriver();					// close port and destroy driver
  CoUninitialize();					// Closes the COM library on the current thread
  cleanup_common();					// delete log-file
  //fclose(CfgFile);					// close .ini file
}
//---------------------------------------------------------------------------------
class myClassFactory: public IClassFactory	
{
public:
  LONG RefCount;			// reference counter
  LONG server_count;		// server counter
  CRITICAL_SECTION lk;		// protect RefCount 
  
  myClassFactory(): RefCount(0), server_count(0)	// when creating interface zero all counter
  {  InitializeCriticalSection(&lk);	  }
  ~myClassFactory()
  {  DeleteCriticalSection(&lk); }

// IUnknown realisation
  STDMETHODIMP QueryInterface(REFIID, LPVOID*);
  STDMETHODIMP_(ULONG) AddRef( void);
  STDMETHODIMP_(ULONG) Release( void);
// IClassFactory realisation
  STDMETHODIMP CreateInstance(LPUNKNOWN, REFIID, LPVOID*);
  STDMETHODIMP LockServer(BOOL);
  
  inline LONG getRefCount(void)
  {
    LONG rc;
    EnterCriticalSection(&lk);		// attempt recieve variable whom may be used by another threads
    rc = RefCount;					// rc = client counter
    LeaveCriticalSection(&lk);                                    
    return rc;
  }

  inline int in_use(void)
  {
    int rv;
    EnterCriticalSection(&lk);
    rv = RefCount | server_count;
    LeaveCriticalSection(&lk);
    return rv;
  }

  inline void serverAdd(void)
  {
    InterlockedIncrement(&server_count);	// increment server counter
  }

  inline void serverRemove(void)
  {
    InterlockedDecrement(&server_count);	// decrement server counter
  }

};
//----- IUnknown -------------------------------------------------------------------------
STDMETHODIMP myClassFactory::QueryInterface(REFIID iid, LPVOID* ppInterface)
{
  if (ppInterface == NULL) return E_INVALIDARG;	// pointer to interface missed (NULL)

  if (iid == IID_IUnknown || iid == IID_IClassFactory)	// legal IID
    {
      UL_DEBUG((LOGID, "myClassFactory::QueryInterface() Ok"));	// write to log
      *ppInterface = this;		// interface succesfully returned
      AddRef();					// adding reference to interface
      return S_OK;				// return succesfully
    }
  UL_ERROR((LOGID, "myClassFactory::QueryInterface() Failed"));
  *ppInterface = NULL;			// no interface returned
  return E_NOINTERFACE;			// error = No Interface
}

STDMETHODIMP_(ULONG) myClassFactory::AddRef(void)	// new client was connected
{
  ULONG rv;
  EnterCriticalSection(&lk);
  rv = (ULONG)++RefCount;							// increment counter of client
  LeaveCriticalSection(&lk);                                    
  UL_DEBUG((LOGID, "myClassFactory::AddRef(%ld)", rv));	// write to log number
  return rv;
}

STDMETHODIMP_(ULONG) myClassFactory::Release(void)	// client has been disconnected
{
  ULONG rv;
  EnterCriticalSection(&lk);
  rv = (ULONG)--RefCount;							// decrement client counter
  LeaveCriticalSection(&lk);
  UL_DEBUG((LOGID, "myClassFactory::Release(%d)", rv));	// write to log number
  return rv;
}

//----- IClassFactory ----------------------------------------------------------------------
STDMETHODIMP myClassFactory::LockServer(BOOL fLock)
{
  if (fLock)
    AddRef();
  else
    Release();
  UL_DEBUG((LOGID, "myClassFactory::LockServer(%d)", fLock)); 
  return S_OK;
}

STDMETHODIMP myClassFactory::CreateInstance(LPUNKNOWN pUnkOuter, REFIID riid, LPVOID* ppvObject)
{
  if (pUnkOuter != NULL)
    return CLASS_E_NOAGGREGATION; // Aggregation is not supported by this code

  IUnknown *server = 0;

  AddRef(); /* for a_server_finished() */
  if (loClientCreate(my_service, (loClient**)&server, 0, &vendor, a_server_finished, this))
    {
      UL_DEBUG((LOGID, "myClassFactory::loCreateClient() failed"));
      Release();
      return E_OUTOFMEMORY;	
    }
  serverAdd();

  HRESULT hr = server->QueryInterface(riid, ppvObject);
  if (FAILED(hr))
    {
      UL_DEBUG((LOGID, "myClassFactory::loClient QueryInterface() failed"));
    }
  else
    {
      loSetState(my_service, (loClient*)server, loOP_OPERATE, OPCstatus, 0);
      UL_DEBUG((LOGID, "myClassFactory::server_count = %ld", server_count));
    }
  server->Release();
  return hr;
}

static myClassFactory my_CF;
static void a_server_finished(void *a, loService *b, loClient *c)
{
  my_CF.serverRemove();						
  if (a) ((myClassFactory*)a)->Release();
  UL_DEBUG((LOGID, "a_server_finished(%lu)", my_CF.server_count));
}

//---------------------------------------------------------------------------------
int APIENTRY WinMain(HINSTANCE hInstance,HINSTANCE hPrevInstance,LPSTR lpCmdLine,int nCmdShow)
{
  static char *argv[3] = { "dummy.exe", NULL, NULL };	// defaults arguments
  argv[1] = lpCmdLine;									// comandline - progs keys
  return mymain(hInstance, 2, argv);
}

int main(int argc, char *argv[])
{
  return mymain(GetModuleHandle(NULL), argc, argv);
}

int mymain(HINSTANCE hInstance, int argc, char *argv[]) 
{
const char eClsidName [] = ECL_SID;				// desription 
const char eProgID [] = ECL_SID;				// name
DWORD objid;									// fully qualified path for the specified module
char *cp;
objid=::GetModuleFileName(NULL, argv0, sizeof(argv0));	// function retrieves the fully qualified path for the specified module
if(objid==0 || objid+50 > sizeof(argv0)) return 0;		// not in border


init_common();									// create log-file
if(NULL==(cp = setlocale(LC_ALL, ".1251")))		// sets all categories, returning only the string cp-1251
	{ 
	UL_ERROR((LOGID, "setlocale() - Can't set 1251 code page"));	// in bad case write error in log
	cleanup_common();							// delete log-file
    return 0;
	}
cp = argv[1];		
if(cp)						// check keys of command line 
	{
    int finish = 1;			// flag of comlection
    if (strstr(cp, "/r"))	//	attempt registred server
		{
	     if (loServerRegister(&UHSID_hOPCserverExe, eProgID, eClsidName, argv0, 0)) 
			{ show_error("Registration Failed");
			  UL_ERROR((LOGID, "Registration <%s> <%s> Failed", eProgID, argv0));  } 
		 else 
			{ show_msg("Hart OPC Registration Ok");
			 UL_INFO((LOGID, "Registration <%s> <%s> Ok", eProgID, argv0));		}
		} 
	else 
		if (strstr(cp, "/u")) 
			{
			 if (loServerUnregister(&UHSID_hOPCserverExe, eClsidName)) 
				{ show_error("UnRegistration Failed");
				 UL_ERROR((LOGID, "UnReg <%s> <%s> Failed", eClsidName, argv0)); } 
			 else 
				{ show_msg("Hart OPC Server Unregistered");
				 UL_INFO((LOGID, "UnReg <%s> <%s> Ok", eClsidName, argv0));		}
			} 
		else  // only /r and /u options
			if (strstr(cp, "/?")) 
				 show_msg("Use: \nKey /r to register server.\nKey /u to unregister server.\nKey /? to show this help.");
				 else
					{
					 UL_WARNING((LOGID, "Ignore unknown option <%s>", cp));
					 finish = 0;		// nehren delat
					}
		if (finish) {      cleanup_common();      return 0;    } 
	}
if ((CfgFile = fopen(CFG_FILE, "r+")) == NULL )	
	{
	 show_error("Error open .ini file");
	 UL_ERROR((LOGID, "Error open .ini file"));	// in bad case write error in log
	 return 0;
	}

if (FAILED(CoInitializeEx(NULL, COINIT_MULTITHREADED))) {	// Initializes the COM library for use by the calling thread
    UL_ERROR((LOGID, "CoInitializeEx() failed. Exiting..."));
    cleanup_common();	// close log-file
    return 0;
  }
UL_INFO((LOGID, "CoInitializeEx() Ok...."));	// write to log

devp = new Hart();

if (InitDriver()) {		// open and set com-port
    CoUninitialize();	// Closes the COM library on the current thread
    cleanup_common();	// close log-file
    return 0;
  }

UL_INFO((LOGID, "InitDriver() Ok...."));	// write to log

if (FAILED(CoRegisterClassObject(UHSID_hOPCserverExe, &my_CF, 
				   CLSCTX_LOCAL_SERVER|CLSCTX_REMOTE_SERVER|CLSCTX_INPROC_SERVER, 
				   REGCLS_MULTIPLEUSE, &objid)))
    { UL_ERROR((LOGID, "CoRegisterClassObject() failed. Exiting..."));
      cleanup_all(objid);		// close comport and unload all librares
      return 0; }
UL_INFO((LOGID, "CoRegisterClassObject() Ok...."));	// write to log

Sleep(3000);			
my_CF.Release();		// avoid locking by CoRegisterClassObject() 

if (OPCstatus!=OPC_STATUS_RUNNING)	// ???? maybe Status changed and OPC not currently running??
	{	while(my_CF.in_use())      Sleep(1000);	// wait
		cleanup_all(objid);
		return 0;	}
addCommToPoll();		// check tags list and list who need
while(my_CF.in_use())	// while server created or client connected
    poll_device();		// polling devices else do nothing (and be nothing)

cleanup_all(objid);		// destroy himself
return 0;
}
//-------------------------------------------------------------------
void poll_device(void)
{
  FILETIME ft;
  int devi, i,j, ecode;
  for (devi = 0, i = 0; devi < devp->idnum; devi++) 
	{
	  UL_DEBUG((LOGID, "Driver poll <%d>", devp->ids[i]));
	  ecode=PollDevice(devp->ids[i]);
		if (ecode)
			{
			dataToTag (devp->ids[i]);
			UL_DEBUG((LOGID, "Copy data to tag success"));
			//time(&devp->mtime);
			//timetToFileTime(devp->mtime, &ft);
			}
  		else
			{
			for (j=0; j<devp->cv_size; j++)
			    devp->cv_status[j] = FALSE;
			GetSystemTimeAsFileTime(&ft);
		}	 
	 EnterCriticalSection(&lk_values);
	 int ci;
	 for (ci = 0; ci < devp->cv_size; ci++, i++) 
		{
	      WCHAR buf[64];
		  LCID lcid = MAKELCID(0x0409, SORT_DEFAULT); // This macro creates a locale identifier from a language identifier. Specifies how dates, times, and currencies are formatted
 	 	  MultiByteToWideChar(CP_ACP,	 // ANSI code page
									  0, // flags
						   devp->cv[ci], // points to the character string to be converted
				 strlen(devp->cv[ci])+1, // size in bytes of the string pointed to 
									buf, // Points to a buffer that receives the translated string
			  sizeof(buf)/sizeof(buf[0])); // function maps a character string to a wide-character (Unicode) string
// 	  	  UL_DEBUG((LOGID, "set tag <i,status,value>=<%d,%d,%s,%s>",i, devp->cv_status[ci], devp->cv[ci],buf));
	
		  if (devp->cv_status[ci]) 
			{	
			  VARTYPE tvVt = tv[i].tvValue.vt;
			  VariantClear(&tv[i].tvValue);
			  switch (tvVt) 
					{
				 	  case VT_I2:
								short vi2;
								vi2 = *devp->cv[ci];
								VarI2FromStr(buf, lcid, 0, &vi2);
								V_I2(&tv[i].tvValue) = vi2;
								break;
				 	  case VT_I4:
								ULONG vi4;								
								vi4=c24bitto32bit (devp->cv[ci]);
								V_I4(&tv[i].tvValue) = vi4;
								break;
					  case VT_R4:
						  	    float vr4;
								vr4= (float) cIEEE754toFloatC (devp->cv[ci]);
								V_R4(&tv[i].tvValue) = vr4;
								break;
					  case VT_UI1:
						  	    unsigned char vui1;
								vui1 = *devp->cv[ci];
							    V_UI1(&tv[i].tvValue) = vui1;
								break;
					  case VT_DATE:
						  	    DATE date; 
							    VarDateFromStr(buf, lcid, 0, &date);
								V_DATE(&tv[i].tvValue) = date;
								break;
					  case VT_BSTR:
					  default:
//								char mess;
//								mess = *devp->cv[ci];
//								V_BSTR(&tv[i].tvValue) = SysAllocString((unsigned short)devp->cv[ci]);
//								FromBase64String(devp->cv[ci],strlen(devp->cv[ci])+1,buf,j);
								//char* Base64toString (char *Data,int lenght)
//				 		 		MultiByteToWideChar(CP_ACP,0,Base64toString (devp->cv[ci],strlen(devp->cv[ci])),strlen(Base64toString (devp->cv[ci],strlen(devp->cv[ci])))+1, buf, sizeof(buf)/sizeof(buf[0]));
								V_BSTR(&tv[i].tvValue) = SysAllocString(buf);
					}
				V_VT(&tv[i].tvValue) = tvVt;
		  	    tv[i].tvState.tsQuality = OPC_QUALITY_GOOD;
			}
			if (ecode == 0)
					tv[i].tvState.tsQuality = OPC_QUALITY_UNCERTAIN;
			if (ecode == 2)
					tv[i].tvState.tsQuality = OPC_QUALITY_DEVICE_FAILURE;
			tv[i].tvState.tsTime = ft;
		}
	 loCacheUpdate(my_service, tTotal, tv, 0);
     LeaveCriticalSection(&lk_values);
	}
  UL_DEBUG((LOGID, "sleep"));
  Sleep(1000);
}
//-------------------------------------------------------------------
loTrid ReadTags(const loCaller *, unsigned  count, loTagPair taglist[],
		VARIANT   values[],	WORD      qualities[],	FILETIME  stamps[],
		HRESULT   errs[],	HRESULT  *master_err,	HRESULT  *master_qual,
		const VARTYPE vtype[],	LCID lcid)
{  return loDR_STORED; }
//-------------------------------------------------------------------
int WriteTags(const loCaller *ca,
              unsigned count, loTagPair taglist[],
              VARIANT values[], HRESULT error[], HRESULT *master, LCID lcid)
{  
 unsigned ii,ci,devi; int i;
 char cmdData[30];	// data to save massive
 char *ppbuf = cmdData;
 VARIANT v;				// input data - variant type
 char ldm;				
 struct lconv *lcp;		// Contains formatting rules for numeric values in different countries/regions
 lcp = localeconv();	// Gets detailed information on locale settings.	
 ldm = *(lcp->decimal_point);	// decimal point (i nah ona nujna?)
 VariantInit(&v);				// Init variant type
 UL_TRACE((LOGID, "WriteTags (%d) invoked", count));	
 EnterCriticalSection(&lk_values);	
 for(ii = 0; ii < count; ii++) 
	{
      HRESULT hr = 0;
	  loTagId clean = 0;
      cmdData[0] = '\0';
      cmdData[HARTCOM_DATALEN_MAX] = '\0';
      UL_TRACE((LOGID,  "WriteTags(Rt=%u Ti=%u)", taglist[ii].tpRt, taglist[ii].tpTi));	  
      i = (unsigned)taglist[ii].tpRt - 1;
      ci = i % devp->cv_size;
      devi = i / devp->cv_size;
      if (!taglist[ii].tpTi || !taglist[ii].tpRt || i >= tTotal)
			continue;
      VARTYPE tvVt = tv[i].tvValue.vt;
      hr = VariantChangeType(&v, &values[ii], 0, tvVt);
      if (hr == S_OK) 
		{
			switch (tvVt) 
				{
				 case VT_I2: _snprintf(cmdData, HARTCOM_DATALEN_MAX, "%d", v.iVal);
							 break; // Write formatted data to a string
				 case VT_R4: UL_TRACE((LOGID, "Number input (%f)",v.fltVal));
							 _snprintf(cmdData, HARTCOM_DATALEN_MAX, "%f", v.fltVal);
							 ppbuf = cFloatCtoIEEE754 (v.fltVal);
							 strcpy(cmdData, ppbuf);
							 //if (ldm != '.' && (dm = strchr(cmdData, ldm)))  *dm = '.';
							 break;
				 case VT_UI1:_snprintf(cmdData, HARTCOM_DATALEN_MAX, "%c", v.bVal);
							 break;
				 case VT_BSTR:
							 WideCharToMultiByte(CP_ACP,0,v.bstrVal,-1,cmdData,HARTCOM_DATALEN_MAX,NULL, NULL);
							 if (strstr(tn[i],"Message")) 
							 {
								 UL_TRACE((LOGID, "WriteMessage(%s)",cmdData));
								 ppbuf=StringtoBase64 (cmdData,strlen(cmdData));
								 strcpy(cmdData, ppbuf);
								 UL_TRACE((LOGID, "WriteMessage(%s)",cmdData));
							 } break;
				 default:	  
							 WideCharToMultiByte(CP_ACP,0,v.bstrVal,-1,cmdData,HARTCOM_DATALEN_MAX,NULL, NULL);
				}
		 UL_TRACE((LOGID, "%!l WriteTags(Rt=%u Ti=%u %s %s)",hr, taglist[ii].tpRt, taglist[ii].tpTi, tn[i], cmdData));
 		 hr = WriteDevice(devp->ids[devi], devp->cv_cmdid[ci], cmdData);
	  }
       VariantClear(&v);
	   if (S_OK != hr) 
			{
			 *master = S_FALSE;
			 error[ii] = hr;
			 UL_TRACE((LOGID, "Write failed"));
			}
	   UL_TRACE((LOGID, "Write success"));
       taglist[ii].tpTi = clean; 
  }
 LeaveCriticalSection(&lk_values); 
 return loDW_TOCACHE; 
}
//-------------------------------------------------------------------
void activation_monitor(const loCaller *ca, int count, loTagPair *til)
{}
//-------------------------------------------------------------------
CHAR Device[16][300];			// id устройств
CHAR DeviceD[16][300];		// data с устройств
int tcount=0;
struct DeviceData 
	{
		char manufacturerId;			// manufacturer ID		
		unsigned char devicetypeId;		// device type ID
		unsigned char preambleNo;		// number of preambles
		unsigned char uncomrevNo;		// universal command rev.No
		unsigned char dccrevNo;			// device-specific command rev.No
		unsigned char softRev;			// software revision
		unsigned char hardRev;			// hardware revision
		unsigned char deviceId[3];		// device identification
		unsigned char hUnitCodePV;		// HART unit code of PV
		unsigned char PV[4];			// primary process variable
		unsigned char hUnitCodeSV;		// HART unit code of SV
		unsigned char SV[4];			// secondary process variable
		unsigned char hUnitCodeTV;		// HART unit code of TV
		unsigned char TV[4];			// third process variable
		unsigned char hUnitCodeFV;		// HART unit code of FV
		unsigned char FV[4];			// fourth process variable
		unsigned char hAC[4];			// actual current of PV
		unsigned char hPrecent[4];		// percentage of measuring range
		unsigned char message[25];		// user message
	}; 	
struct DeviceData DData[16];
unsigned char DeviceDataBuffer[16][15][30];
unsigned int mBuf[100];
//----------------------------------------------------------------------------
void addCommToPoll()
{
int	max_com=0,i,j,flajok;	mBuf[0]=9588;
for (i=0;i<sizeof(HartCommU)/sizeof(HartCom);i++)
{	if (HartCommU[i].getCmd!=-1) 
	{flajok=0;
		for (j=0;j<=tcount;j++)
			if (HartCommU[i].getCmd==mBuf[j])	flajok=1;
	if (flajok==0) { mBuf[tcount]=HartCommU[i].getCmd; 
	UL_DEBUG((LOGID, "Add Hart command %d to poll. Total %d command added.",mBuf[tcount], tcount)); tcount++;}
	}
}}

void dataToTag(int device)
{
int	max_com=0,i,k; 
unsigned int j;	
char buf[50];	char buf1[50];
char *pbuf=buf; 
*pbuf = '\0';
for (i=0; i<devp->cv_size; i++)			
	{
	 if (HartCommU[i].scan_it)
	 {
	 for (j=0; j<HartCommU[i].num; j++)
		{
		 for (k=0;k<tcount;k++) 
			 if (mBuf[k]==HartCommU[i].getCmd) 
				{ 
				 UL_DEBUG((LOGID, "tag: %s / com: %d / sym: %d",HartCommU[i].name,mBuf[k],DeviceDataBuffer[device][mBuf[k]][j+2+HartCommU[i].start])); 
				 break;
				}
		 buf[j]=DeviceDataBuffer[device][mBuf[k]][j+2+HartCommU[i].start];
		}
	 buf[j] = '\0';	
	 if (HartCommU[i].name=="Date") { pbuf = cDharttoDate (buf); strcpy(devp->cv[i], pbuf);}
	 else
	 {
	 if ((HartCommU[i].name=="Message")||(HartCommU[i].name=="TAG")||(HartCommU[i].name=="Descriptor")) 
		{
		 pbuf = Base64toString (buf,strlen(buf));
		 strcpy(devp->cv[i], pbuf);}
	 else 
		{
		 BOOL tz=FALSE;
		 for (j=0;j<HartCommU[i].num;j++) 
			 if (buf[j]==0) 
				 if (buf[j+1]==0)
					tz=TRUE;
		 if (tz)
			{
			 UL_INFO((LOGID, "two zero found"));
			 for (j=0;j<HartCommU[i].num;j++) 
				*(devp->cv[i]+j)=buf[j];
			 *(devp->cv[i]+j)=0;
			}
		 else					
			strcpy (devp->cv[i],buf);
		}
	 }
	 devp->cv_status[i] = TRUE;
	 }
	 else 
		{
		 if (HartCommU[i].name=="Manufacturer Name")
			{	
				_itoa((int)DData[device].manufacturerId,buf,16);
				sprintf(buf,"%s",buf);
				pbuf = ReadParam ("Manufacturers",buf);
				strcpy(devp->cv[i], pbuf);
				devp->cv_status[i] = TRUE;
			}
		 if (HartCommU[i].name=="Device Name")
			{
				_itoa((int)DData[device].manufacturerId,buf1,16);
				_itoa((int)DData[device].devicetypeId,buf,16);
				sprintf(buf1,"%s%s",buf1,buf);
				pbuf = ReadParam ("Models",buf1);
				strcpy(devp->cv[i], pbuf);
				devp->cv_status[i] = TRUE;
			}
		 if (HartCommU[i].name=="PV unit")
			{	
				_itoa((int)DData[device].hUnitCodePV,buf,16);
				sprintf(buf,"%s",buf);
				pbuf = ReadParam ("Units",buf);
				strcpy(devp->cv[i], pbuf);				
				devp->cv_status[i] = TRUE;
			}
 		 if (HartCommU[i].name=="SV unit")
			{	_itoa((int)DData[device].hUnitCodeSV,buf,16);
				sprintf(buf,"%s",buf);
				pbuf = ReadParam ("Units",buf);
				strcpy(devp->cv[i], pbuf);				
				devp->cv_status[i] = TRUE;
			}
		 if (HartCommU[i].name=="TV unit")
			{	_itoa((int)DData[device].hUnitCodeTV,buf,16);
				sprintf(buf,"%s",buf);
				pbuf = ReadParam ("Units",buf);
				strcpy(devp->cv[i], pbuf);				
				devp->cv_status[i] = TRUE;
			}
		 if (HartCommU[i].name=="QV unit")
			{	_itoa((int)DData[device].hUnitCodeFV,buf,16);
				sprintf(buf,"%s",buf);
				pbuf = ReadParam ("Units",buf);
				strcpy(devp->cv[i], pbuf);				
				devp->cv_status[i] = TRUE;
			}
		}
//	 UL_DEBUG((LOGID, "Copy data to tag %s. Num: %d. Start: %d.",devp->cv[i],HartCommU[i].num,HartCommU[i].start));
	}
}
//-----------------------------------------------------------------------------------
// data - massive to send 
HRESULT WriteDevice(int device,const unsigned cmdnum,LPSTR data)
{
	int nump,cnt_false,cnt;
	unsigned int k,j,i;
	unsigned char Out[45];
	unsigned char sBuf1[40];
	const int sBuf[] = {0xff,0xff,0xff,0xff,0xff,0x2,0x80,0x0,0x0,0x82};
	unsigned char *Outt = Out,*Int = sBuf1; 
	DWORD dwStoredFlags;
	COMSTAT comstat;
	dwStoredFlags = EV_RXCHAR | EV_TXEMPTY | EV_RXFLAG;	
	UL_DEBUG((LOGID, "command send (write) %d",cmdnum));
	port.SetMask (dwStoredFlags);
	port.GetStatus (comstat);	

	for (nump=0;nump<=1;nump++)
	{	
	port.SetRTS ();	port.ClearDTR ();
	for (i=0;i<=13;i++) 	Out[i] = (char) sBuf[i];
	Out[5]=0x82; 
	Out[6]=DData[device].manufacturerId; Out[7]=DData[device].devicetypeId;  
	Out[8]=DData[device].deviceId[0]; Out[9]=DData[device].deviceId[1]; 
	Out[10]=DData[device].deviceId[2]; Out[11]=HartCommU[cmdnum].setCmd;
	Out[12]=HartCommU[cmdnum].num; 
	for (j=13;j<13+HartCommU[cmdnum].num;j++)  Out[j] = data[j-13];
 	
	Out[j] = 0;	for (k=5;k<j;k++) Out[j]=Out[j]^Out[k]; 
	j=Out[12]+15;
	for (i=0;i<=j;i++) 
	{	
	 port.Write(Outt+i, 1);	port.WaitEvent (dwStoredFlags);
	 UL_DEBUG((LOGID, "send sym (write): %d = %d / all: %d",i,Out[i],Out[12]+1));	 
	}
	port.ClearRTS ();	port.SetDTR ();	
	cnt_false=0;
	for (cnt=0;cnt<40;cnt++)
	{
	if (port.Read(Int+cnt, 1) == FALSE)
		{cnt_false++; if (cnt_false>2) break;}
	else cnt_false=0;
	UL_DEBUG((LOGID, "recieve sym (write): %d = %d.",cnt,sBuf1[cnt]));
	port.GetStatus (comstat);
	}
//-----------------------------------------------------------------------------------
	BOOL bcFF_OK = FALSE; BOOL bcFF_06 = FALSE; int chff=0;
	for (cnt=0;cnt<26;cnt++)
		{ 
		 if (sBuf1[cnt]==0xff) {chff++; if (chff>3) bcFF_OK = TRUE;} else chff=0;
		 if (bcFF_OK) 
			 if (sBuf1[cnt]==0x86)								// ответ от подчиненного???
			 if (sBuf1[cnt+1]==DData[device].manufacturerId)	// да еще и тот адрес
				if (sBuf1[cnt+2]==DData[device].devicetypeId)	// да еще и на команду идентификации!
						{ bcFF_06 = TRUE;  break;}
		}
	if (bcFF_06)
		  return S_OK;
	}
	return E_FAIL;
}
//-----------------------------------------------------------------------------------
UINT PollDevice(int device)
{
	int i=0,cnt=0,c0m=0, nump, chff=0, num_bytes=0, startid=0, cnt_false=0;
	const int sBuf[] = {0xff,0xff,0xff,0xff,0xff,0x2,0x80,0x0,0x0,0x82};
	//Device[device]="";		// id устройств
	//DeviceD[device]="";		// data с устройств
	unsigned char sBuf1[40];
	unsigned char Out[15];
	unsigned char DId[80];
	unsigned char Data[30];
	unsigned char *Outt = Out,*DeviceId = DId,*Int = sBuf1;
	char buffer[12];
	COMSTAT comstat;
	DWORD dwStoredFlags;
	BOOL bcFF_OK = FALSE;
	BOOL bcFF_06 = FALSE;
	dwStoredFlags = EV_RXCHAR | EV_TXEMPTY | EV_RXFLAG;	
	port.SetMask (dwStoredFlags);
	port.GetStatus (comstat);	
//-----------------------------------------------------------------------------------	
	for (nump=0;nump<=1;nump++)			//!!! (1)
	{
	port.SetRTS ();
	if (DTRHIGH) port.SetDTR (); else port.ClearDTR ();
	for (i=0;i<=11;i++) 	Out[i] = (char) sBuf[i]; // 5 первых ff + стандарт команды
	Out[6]=0x80+device; Out[7]=0; Out[9]=Out[5]^Out[6]^Out[7]^Out[8];
	for (i=0;i<=11;i++) 
	{		port.Write(Outt+i, 1);	port.WaitEvent (dwStoredFlags);	}	
	port.ClearRTS ();	port.SetDTR ();
	Int = sBuf1;	cnt_false=0;
	for (cnt=0;cnt<26;cnt++)
	{
	if (port.Read(Int+cnt, 1) == FALSE)
		{cnt_false++; if (cnt_false>1) break;}	//!!! (5)
	else cnt_false=0;
	port.GetStatus (comstat);
	}
	}
//-----------------------------------------------------------------------------------
	for (cnt=0;cnt<36;cnt++)
		{ 
		 if (sBuf1[cnt]==0xff) {chff++; if (chff>3) bcFF_OK = TRUE;} else chff=0;		  
		 if (bcFF_OK) 
			 if (sBuf1[cnt]==0x6)				// ответ от подчиненного???
			 if (sBuf1[cnt+1]==(0x80+device))	// да еще и тот адрес
				if (sBuf1[cnt+2]==0)			// да еще и на команду идентификации!
								{ bcFF_06 = TRUE; num_bytes = sBuf1[cnt+3]; startid=cnt+4;}
		}
//-----------------------------------------------------------------------------------				
	if (bcFF_06) 
	{		
	DId[device*5] = 0x80+sBuf1[startid+3];
	DId[device*5+1] = sBuf1[startid+4];
	DId[device*5+2] = sBuf1[startid+11];
	DId[device*5+3] = sBuf1[startid+12];
	DId[device*5+4] = sBuf1[startid+13];

	DData[device].manufacturerId = sBuf1[startid+3];
	DData[device].devicetypeId = sBuf1[startid+4];
	DData[device].preambleNo = sBuf1[startid+5];
	DData[device].uncomrevNo = sBuf1[startid+6];
	DData[device].dccrevNo = sBuf1[startid+7];
	DData[device].softRev = sBuf1[startid+8];
	DData[device].hardRev = sBuf1[startid+9];
	DData[device].deviceId[0] = sBuf1[startid+11];
	DData[device].deviceId[1] = sBuf1[startid+12];
	DData[device].deviceId[2] = sBuf1[startid+13];

	for (c0m=0;c0m<tcount;c0m++) for (nump=0;nump<=1;nump++)
//	for (c0m=144;c0m<146;c0m++) for (nump=0;nump<=1;nump++)
	{	
	 UL_DEBUG((LOGID, "command request %d",mBuf[c0m]));
//-----------------------------------------------------------------------------------
	 port.SetRTS ();
	 if (DTRHIGH) port.SetDTR (); else port.ClearDTR ();	
	for (i=0;i<=13;i++) 	Out[i] = (char) sBuf[i];	
	Out[5]=0x82; 
	Out[6]=DId[device*5]; Out[7]=DId[device*5+1];  Out[8]=DId[device*5+2]; Out[9]=DId[device*5+3]; 
	Out[10]=DId[device*5+4]; Out[11]=mBuf[c0m];  Out[12]=0x0; 
//	Out[10]=DId[device*5+4]; Out[11]=c0m;  Out[12]=0x0; 
	Out[13]=Out[5]^Out[6]^Out[7]^Out[8]^Out[9]^Out[10]^Out[11]^Out[12]^Out[13]; 
	for (i=0;i<=15;i++) 
	{	
	port.Write(Outt+i, 1);
	port.WaitEvent (dwStoredFlags);
	}
	port.ClearRTS ();	port.SetDTR ();	
	cnt_false=0;
	for (cnt=0;cnt<40;cnt++)
	{
	if (port.Read(Int+cnt,1)==FALSE)
		{{cnt_false++; cnt--;} if (cnt_false>1) break;}		//!!! (2)
	else cnt_false=0;
	port.GetStatus (comstat);	
	}
//-----------------------------------------------------------------------------------
	BOOL bcFF_OK = FALSE;
	BOOL bcFF_06 = FALSE; chff=0;
	for (cnt=0;cnt<26;cnt++)
		{ 
		 if (sBuf1[cnt]==0xff) {chff++; if (chff>3) bcFF_OK = TRUE;} else chff=0;
		 if (bcFF_OK) 
			 if (sBuf1[cnt]==0x86)						// ответ от подчиненного???
			 if (sBuf1[cnt+1]==DId[device*5])			// да еще и тот адрес
				if (sBuf1[cnt+2]==DId[device*5+1])		// да еще и на команду идентификации!
								{ bcFF_06 = TRUE; num_bytes = sBuf1[cnt+7]; startid=cnt+8; break;}
		}
	if (bcFF_06)
		{
		unsigned char ks=0;
		for (i=startid-8;i<startid+num_bytes;i++)
			 ks=ks^sBuf1[i];
		if (ks!=sBuf1[i])	{ bcFF_06=FALSE; break;}
//			UL_DEBUG((LOGID, "ks ok. ks:%d // need:%d",ks,sBuf1[i]));
//		else
//			UL_DEBUG((LOGID, "ks error. ks:%d // need:%d",ks,sBuf1[i]));
		for (cnt=0;cnt<30;cnt++) Data[cnt]=0;
		for (cnt=0;cnt<num_bytes;cnt++)	  
			{ 
				Data[cnt]=sBuf1[startid+cnt]; 
				DeviceDataBuffer[device][mBuf[c0m]][cnt] = sBuf1[startid+cnt];
//				UL_DEBUG((LOGID, "sym: %d = %d.",cnt,DeviceDataBuffer[device][mBuf[c0m]][cnt]));
			}
		for (int k=0;k<devp->cv_size;k++)  
			if (HartCommU[k].getCmd==mBuf[c0m])
			  devp->cv_status[k]=TRUE;

		if (mBuf[c0m]==0x1)
		{
		  DData[device].hUnitCodePV = Data[2];
		  DData[device].PV[0] = Data[3]; 
		  DData[device].PV[1] = Data[4];
		  DData[device].PV[2] = Data[5]; 
		  DData[device].PV[3] = Data[6];
		}
		if (mBuf[c0m]==0x2)
		{
		  DData[device].hAC[0] = Data[2]; 
		  DData[device].hAC[1] = Data[3];
		  DData[device].hAC[2] = Data[4]; 
		  DData[device].hAC[3] = Data[5];
		  DData[device].hPrecent[0] = Data[6]; 
		  DData[device].hPrecent[1] = Data[7];
		  DData[device].hPrecent[2] = Data[8]; 
		  DData[device].hPrecent[3] = Data[9];
		}
		if (mBuf[c0m]==0x3)
		{
		  DData[device].hAC[0] = Data[2]; 
		  DData[device].hAC[1] = Data[3];
		  DData[device].hAC[2] = Data[4]; 
		  DData[device].hAC[3] = Data[5];
		  DData[device].hUnitCodePV = Data[6];
		  DData[device].PV[0] = Data[7]; 
		  DData[device].PV[1] = Data[8];
		  DData[device].PV[2] = Data[9]; 
		  DData[device].PV[3] = Data[10];
		  DData[device].hUnitCodeSV = Data[11];
		  DData[device].SV[0] = Data[12]; 
		  DData[device].SV[1] = Data[13];
		  DData[device].SV[2] = Data[14]; 
		  DData[device].SV[3] = Data[15];
		  DData[device].hUnitCodeTV = Data[16];
		  DData[device].TV[0] = Data[17]; 
		  DData[device].TV[1] = Data[18];
		  DData[device].TV[2] = Data[19]; 
		  DData[device].TV[3] = Data[20];
		  DData[device].hUnitCodeFV = Data[21];
		  DData[device].FV[0] = Data[22]; 
		  DData[device].FV[1] = Data[23];
		  DData[device].FV[2] = Data[24]; 
		  DData[device].FV[3] = Data[25];
		}
		if (mBuf[c0m]==0xb)
			for (cnt=0;cnt<num_bytes;cnt++)	  
				DData[device].message[cnt]=Data[cnt+2];
	}
	if (bcFF_06) nump=1;
	else 
		{
		 for (int k=0;k<devp->cv_size;k++)
			if (HartCommU[k].getCmd==mBuf[c0m])
			  devp->cv_status[k]=FALSE;
		}
//-----------------------------------------------------------------------------------
	}
	if (bcFF_06) 
		{
			 switch (DData[device].manufacturerId)
				{
				case 17: sprintf_s (Device[device],15," Endress+Hauser. "); break;
				default: sprintf_s(Device[device],15 , " Udentifined. ");
				}

				switch (DData[device].devicetypeId)
				{				
				case 14: sprintf_s(Device[device],15 , " Cerabar M. "); break;
				case 80: sprintf_s(Device[device],15 , " Promass80. "); break;
				case 83: sprintf_s(Device[device],15 , " Promass40. "); break;
				case 200: sprintf_s(Device[device],15 , " TMT182. "); break;
				case 201: sprintf_s(Device[device],15 , " TMT122. "); break;
				case 202: sprintf_s(Device[device],15 , " TMT162. "); break;
				default: sprintf_s(Device[device],15 , " Udentifined. ");
				}
			 sprintf_s(Device[device],15 , " id: "); 
			 cnt = DData[device].deviceId[0] * 256 * 256 + DData[device].deviceId[1] * 256 + DData[device].deviceId[2];
			 sprintf_s(Device[device],15 , _itoa (cnt,buffer,10));
			 sprintf_s(Device[device],15 , " com.rev.No: ");
			 sprintf_s(Device[device],15 , _itoa (DData[device].uncomrevNo,buffer,10));
 			 sprintf_s(Device[device],15 , " dcc.rev.No: ");
			 sprintf_s(Device[device],15 , _itoa (DData[device].dccrevNo,buffer,10));
  			 sprintf_s(Device[device],15 , " soft.rev.No: ");
			 sprintf_s(Device[device],15 , _itoa (DData[device].softRev,buffer,10));
  			 sprintf_s(Device[device],15 , " hard.rev: ");
			 sprintf_s(Device[device],15 , _itoa (DData[device].hardRev,buffer,10));

			 sprintf_s(DeviceD[device],10 , "PV: ");
			 sprintf_s(DeviceD[device],10 , _gcvt (cIEEE754toFloat(DData[device].PV),8,buffer));
		     switch (DData[device].hUnitCodePV)
				{
				case 32: sprintf_s(DeviceD[device],10 , " (°C) "); break;
				case 36: sprintf_s(DeviceD[device],10 , " (mV) "); break;
				case 96: sprintf_s(DeviceD[device],10 , " (kg/l) "); break;
				case 61: sprintf_s(DeviceD[device],10 , " (kg) "); break;
				case 65: sprintf_s(DeviceD[device],10 , " (LTon) "); break;
				case 75: sprintf_s(DeviceD[device],10 , " (kg/h) "); break;
				default: sprintf_s(DeviceD[device],10 , " (h/z) ");
				}
			 sprintf_s(DeviceD[device],10 , " mV: ");
			 sprintf_s(DeviceD[device],10 , _gcvt (cIEEE754toFloat(DData[device].hAC),8,buffer));
 			 sprintf_s(DeviceD[device] ,10, " (");
			 sprintf_s(DeviceD[device],10 , _gcvt (cIEEE754toFloat(DData[device].hPrecent),8,buffer));
			 sprintf_s(DeviceD[device],10 , "%) ");

			 sprintf_s(DeviceD[device],10 , " SV: ");
			 sprintf_s(DeviceD[device],10 , _gcvt (cIEEE754toFloat(DData[device].SV),10,buffer));
			 switch (DData[device].hUnitCodeSV)
				{
				case 32: sprintf_s(DeviceD[device],10 , " (°C) "); break;
				case 36: sprintf_s(DeviceD[device],10 , " (mV) "); break;
				case 96: sprintf_s(DeviceD[device],10 , " (kg/l) "); break;
				case 61: sprintf_s(DeviceD[device],10 , " (kg) "); break;
				case 65: sprintf_s(DeviceD[device],10 , " (LTon) "); break;
				case 75: sprintf_s(DeviceD[device],10 , " (kg/h) "); break;
				default: sprintf_s(DeviceD[device],10 , " (h/z) ");
				}			 
			 sprintf_s(DeviceD[device],10 , " TV: ");
			 sprintf_s(DeviceD[device],10 , _gcvt (cIEEE754toFloat(DData[device].TV),8,buffer));
			 switch (DData[device].hUnitCodeTV)
				{
				case 32: sprintf_s(DeviceD[device],10 , " (°C) "); break;
				case 36: sprintf_s(DeviceD[device],10 , " (mV) "); break;
				case 96: sprintf_s(DeviceD[device],10 , " (kg/l) "); break;
				case 65: sprintf_s(DeviceD[device],10 , " (LTon) "); break;
				case 61: sprintf_s(DeviceD[device],10 , " (kg) "); break;
				case 75: sprintf_s(DeviceD[device],10 , " (kg/h) "); break;
				default: sprintf_s(DeviceD[device],10 , " (h/z) ");
				}			 
			 sprintf_s(DeviceD[device],10 , " FV: ");
			 sprintf_s(DeviceD[device],10 , _gcvt (cIEEE754toFloat(DData[device].FV),8,buffer));
			 switch (DData[device].hUnitCodeFV)
				{
				case 32: sprintf_s(DeviceD[device],10 , " (°C) "); break;
				case 36: sprintf_s(DeviceD[device],10 , " (mV) "); break;
				case 96: sprintf_s(DeviceD[device],10 , " (kg/l) "); break;
				case 65: sprintf_s(DeviceD[device],10 , " (LTon) "); break;
				case 61: sprintf_s(DeviceD[device],10 , " (kg) "); break;
				case 75: sprintf_s(DeviceD[device],10 , " (kg/h) "); break;
				default: sprintf_s(DeviceD[device],10 , " (h/z) ");
				}			 
			//sprintf_s(DeviceD[device],10," Message: "); 
			//for (cnt=0;cnt<26;cnt++) sprintf_s(DeviceD[device],100,DData[device].message[cnt]);
			UL_DEBUG((LOGID, "Device reply a: %s.",Device[device]));
			UL_DEBUG((LOGID, "Device reply b: %s.",DeviceD[device]));
			return 1;
		}
	}
 return 0;
}
//-------------------------------------------------------------------
int init_tags(void)
{
  FILETIME ft;	//  64-bit value representing the number of 100-ns intervals since January 1,1601
  int i,devi;		
  unsigned rights;	// tag type (read/write)
  int ecode;
  GetSystemTimeAsFileTime(&ft);	// retrieves the current system date and time
  EnterCriticalSection(&lk_values);
  for (i=0; i < tTotal; i++)    
	  tn[i] = new char[TAGNAME_LEN];	// reserve memory for massive
  for (devi = 0, i = 0; devi < devp->idnum; devi++) 
  {
    int ci;
    for (ci=0;ci<devp->cv_size; ci++, i++) 
		{
		 int cmdid = devp->cv_cmdid[ci];
		 sprintf(tn[i], "com%d/id%0.2d/%s",comp, devp->ids[devi], HartCommU[cmdid].name);
		 if (HartCommU[cmdid].getCmd!=-1) rights = rights | OPC_READABLE;
	     if (HartCommU[cmdid].setCmd!=-1) rights = rights | OPC_WRITEABLE;
		 VariantInit(&tv[i].tvValue);
		 switch (HartCommU[ci].dtype)
		 {
// loAddRealTag(loService,loTagId *ti,loRealTag rt,const char *tName,int tFlag,
//              unsigned tRight,VARIANT *tValue,int tEUtype, VARIANT *tEUinfo);
			 case VT_I2:
				 V_I2(&tv[i].tvValue) = 0;
				 V_VT(&tv[i].tvValue) = VT_I2;
				 ecode = loAddRealTag_a(my_service, &ti[i], // returned TagId 
			       (loRealTag)(i+1), tn[i],
			       0,rights, &tv[i].tvValue, -99999, 99999); break;
			 case VT_I4:
				 V_I4(&tv[i].tvValue) = 0;
				 V_VT(&tv[i].tvValue) = VT_I4;
				 ecode = loAddRealTag_a(my_service, &ti[i], // returned TagId 
				  (loRealTag)(i+1), tn[i],
			      0,rights, &tv[i].tvValue, -99999, 99999); break;
			 case VT_R4:
				 V_R4(&tv[i].tvValue) = 0.0;
				 V_VT(&tv[i].tvValue) = VT_R4;
				 ecode = loAddRealTag_a(my_service, &ti[i], (loRealTag)(i+1), // != 0
			       tn[i], 0, rights, &tv[i].tvValue, -99999.0, 99999.0); break;
			 case VT_UI1:
				 V_UI1(&tv[i].tvValue) = 0;
				 V_VT(&tv[i].tvValue) = VT_UI1;
				 ecode = loAddRealTag_a(my_service, &ti[i], (loRealTag)(i+1), // != 0
			       tn[i], 0, rights, &tv[i].tvValue, -99, 99); break;
			 case VT_DATE:
				 V_DATE(&tv[i].tvValue) = 0;
				 V_VT(&tv[i].tvValue) = VT_DATE;
				 ecode = loAddRealTag_a(my_service, &ti[i], (loRealTag)(i+1), // != 0
			     tn[i], 0, rights, &tv[i].tvValue, 0, 0); break;
			 default:
				 V_BSTR(&tv[i].tvValue) = SysAllocString(L"");
				 V_VT(&tv[i].tvValue) = VT_BSTR;
				 ecode = loAddRealTag(my_service, &ti[i], (loRealTag)(i+1), 
			     tn[i], 0, rights, &tv[i].tvValue, 0, 0);
		 }      
      tv[i].tvTi = ti[i];
      tv[i].tvState.tsTime = ft;
      tv[i].tvState.tsError = S_OK;
      tv[i].tvState.tsQuality = OPC_QUALITY_NOT_CONNECTED;
	  UL_TRACE((LOGID, "%!e loAddRealTag(%s) = %u", ecode, tn[i], ti[i]));
	}
  } 
  LeaveCriticalSection(&lk_values);
  if(ecode) 
  {
    UL_ERROR((LOGID, "%!e driver_init()=", ecode));
    return -1;
  }
  return 0;
}
//-------------------------------------------------------------------
UINT DestroyDriver()
{
  if (my_service)		
    {
      int ecode = loServiceDestroy(my_service);
      UL_INFO((LOGID, "%!e loServiceDestroy(%p) = ", ecode));	// destroy derver
      DeleteCriticalSection(&lk_values);						// destroy CS
      my_service = 0;		
    }
 port.Close();
 UL_INFO((LOGID, "Close COM-port"));						// write in log
 return	1;
}
//-------------------------------------------------------------------
UINT InitDriver()
{
 loDriver ld;		// structure of driver description
 LONG ecode;		// error code 

 tTotal = 100;		// total tag quantity
 if (my_service) {	
      UL_ERROR((LOGID, "Driver already initialized!"));
      return 0;
  }
 memset(&ld, 0, sizeof(ld));   
 ld.ldRefreshRate =10000;		// polling time 
 ld.ldRefreshRate_min = 7000;	// minimum polling time
 ld.ldWriteTags = WriteTags;	// pointer to function write tag
 ld.ldReadTags = ReadTags;		// pointer to function read tag
 ld.ldSubscribe = activation_monitor;	// callback of tag activity
 ld.ldFlags = loDF_IGNCASE;	// ignore case
 ld.ldBranchSep = '/';					// hierarchial branch separator
 ecode = loServiceCreate(&my_service, &ld, tTotal);		//	creating loService
 UL_TRACE((LOGID, "%!e loServiceCreate()=", ecode));	// write to log returning code
 if (ecode) return 1;									// error to create service	
 InitializeCriticalSection(&lk_values);

			COMMTIMEOUTS timeouts;
			char buf[50]; char *pbuf=buf; 
			pbuf = ReadParam ("Port","COM");			
			comp = atoi(pbuf);
			pbuf = ReadParam ("Port","Speed");			
			speed = atoi(pbuf);
			port.Open(comp,speed,SerialPort::OddParity, 8, SerialPort::OneStopBit, SerialPort::NoFlowControl, FALSE);
			UL_INFO((LOGID, "Opening port COM%d",comp));		// write in log			
			timeouts.ReadIntervalTimeout = 50; 
			timeouts.ReadTotalTimeoutMultiplier = 0; 
			timeouts.ReadTotalTimeoutConstant = 180;	//!!! (180)
			timeouts.WriteTotalTimeoutMultiplier = 0; 
			timeouts.WriteTotalTimeoutConstant = 50; 
			port.SetTimeouts(timeouts);
			UL_INFO((LOGID, "Set COM-port timeouts 50:0:200:0:50"));		// write in log

		UL_INFO((LOGID, "Scan bus"));									// write in log 
		 if (ScanBus()) { 
							UL_INFO((LOGID, "Total %d devices found",devp->idnum)); 
							if (init_tags()) 
									return 1; 
							else return 0;
						}
			else		{ 
							UL_ERROR((LOGID, "No devices found")); 
							return 1; 
						}
}
//----------------------------------------------------------------------------------------
UINT ScanBus()
{
const int sBuf[] = {0xff,0xff,0xff,0xff,0xff,0x2,0x0,0x0,0x0,0x82};	// short frame
unsigned char Out[15],sBuf1[40];
unsigned char *Outt = Out,*Int = sBuf1;
int cnt,i,cnt_false, chff=0;
DWORD dwStoredFlags;
dwStoredFlags = EV_RXCHAR | EV_TXEMPTY | EV_RXFLAG;	
COMSTAT comstat;
port.SetMask (dwStoredFlags);
port.GetStatus (comstat);	
port.Read(Int, 1);
devp->idnum = 0;

for (int adr=0;adr<=HARTCOM_ID_MAX;adr++)
	for (int nump=0;nump<=3;nump++)
		{
		 port.SetRTS ();		 
		 if (nump<2) port.ClearDTR ();
		 if (nump>1) port.SetDTR ();
		 
		 for (i=0;i<=11;i++) 	Out[i] = (char) sBuf[i]; // 5 первых ff + стандарт команды
		 Out[6]=adr; Out[7]=0; Out[9]=Out[5]^Out[6]^Out[7]^Out[8];
		 if (nump==3) Out[7]=0x10;
		 UL_INFO((LOGID, "out: 0x%x 0x%x 0x%x 0x%x 0x%x 0x%x 0x%x 0x%x 0x%x 0x%x 0x%x 0x%x",Out[0],Out[1],Out[2],Out[3],Out[4],Out[5],Out[6],Out[7],Out[8],Out[9],Out[10],Out[11]));
		 for (i=0;i<=11;i++) 
		 {	port.Write(Outt+i, 1);	port.WaitEvent (dwStoredFlags);	}		 
		 port.SetDTR (); port.ClearRTS ();
		 Int = sBuf1;	cnt_false=0;
		 for (cnt=0;cnt<26;cnt++)
			{
			 if (port.Read(Int+cnt, 1) == FALSE)
				{ cnt_false++; if (cnt_false>4) break;}		//!!! (4)
			 else cnt_false=0;
				port.GetStatus (comstat);
			}
//-----------------------------------------------------------------------------------
    if (cnt_false<4) UL_INFO((LOGID, "in: 0x%x 0x%x 0x%x 0x%x 0x%x 0x%x 0x%x 0x%x 0x%x 0x%x 0x%x 0x%x",sBuf1[0],sBuf1[1],sBuf1[2],sBuf1[3],sBuf1[4],sBuf1[5],sBuf1[6],sBuf1[7],sBuf1[8],sBuf1[9],sBuf1[10],sBuf1[11]));
    BOOL bcFF_OK = FALSE;	BOOL bcFF_06 = FALSE;
	for (cnt=0;cnt<36;cnt++)
		{ 
		 if (sBuf1[cnt]==0xff) {chff++; if (chff>3) bcFF_OK = TRUE;} else chff=0;
		 if (bcFF_OK) 
			 if (sBuf1[cnt]==0x6)			// ответ от подчиненного???
			 if (sBuf1[cnt+1]==(0x80+adr))  // да еще и тот адрес
				if (sBuf1[cnt+2]==0)		// да еще и на команду идентификации!
						bcFF_06 = TRUE; 
		}	
    if (bcFF_06) 
		{	
			if (nump>1) DTRHIGH = TRUE; else DTRHIGH = FALSE;
			devp->ids[devp->idnum] = adr; devp->idnum++;
			UL_INFO((LOGID, "Device found on address %d",adr));	// write in log
			break;
		}
	}	 
return devp->idnum;
}
//-----------------------------------------------------------------------------------
// DD MM YEARS+1900 >> DD-MM-YYYY
char* cDharttoDate (char *Data)
{
char buf[13]; char date[3]; char month[3]; char year[5];
char *pbuf=buf;   
*pbuf = '\0'; buf[9]=0x20; buf[10]=0x20; buf[11]=0x20;
//UL_INFO((LOGID, "input date > %i - %i - %i",(int)*(Data),(int)*(Data+1),(int)*(Data+2)));
_itoa (*(Data+2)+1900,year,10);
_itoa (*(Data+1),month,10); 
_itoa (*(Data),date,10);
year[4]='\0';
sprintf(buf,"%s-%s-%s",date,month,year);
//UL_INFO((LOGID, "output date > %s",buf));
return pbuf;
}
//-----------------------------------------------------------------------------------
// SEEEEEEE EMMMMMMM MMMMMMMM MMMMMMMM  << Data
// [0]      [1]      [2]      [3]
//-----------------------------------------------------------------------------------
double cIEEE754toFloat (unsigned char *Data)
{
BOOL sign;
char exp;
double res=0,zn=0.5, tmp;
unsigned char mask;
if (*(Data+0)&0x80) sign=TRUE; else sign=FALSE;
exp = ((*(Data+0)&0x7f)*2+(*(Data+1)&0x80)/0x80)-127;
for (int j=1;j<=3;j++)		// 23 значная мантисса
	{	
	 mask = 0x80;
	 for (int i=0;i<=7;i++)		
		{ 	 
		 if (j==1&&i==0) {res = res+1; mask = mask/2;}
		 else {
		 res = (*(Data+j)&mask)*zn/mask + res;
		 mask = mask/2; zn=zn/2; }
		}
	}
res = res * pow ((double)2,exp);
tmp = 1*pow ((double)10,-15);
if (res<tmp) res=0;
if (sign) res = -res;
return res;
}
//-----------------------------------------------------------------------------------
char* cFloatCtoIEEE754 (float num)
{
char buf[5]={0,0,0,0,'\0'};
char *pbuf=buf;
float pr12,temp=1, addc;
int i=0,j=0,k=0;
if (num<0) buf[0]=(char)0x80;
if (num>=2) 
	{
	 while (num>=2.0)
		{num=num/2; i++;}
	 i=i+127;
	}
if (num<=1)
	{
	 while (num<1)
		{num=num*2; i++;}
	 i=127-i;
	}
buf[0]=(char)(i>>1); 
buf[1]=(i&0x1)*0x80;
pr12=1; addc=1;				// 
for (j=1;j<=23;j++)			// j=2
	{	 
	 addc = addc/2;			// 0.5
	 temp = temp + addc;	// 1.5
	 if (temp>num)			// 
		temp=pr12;
	 else					
		{
		 pr12=temp;			// pr12=1.5
		 i=1+j/8;			// 1
		 k=(int) pow ((double)2,7-j%8);	// 0x20
		 buf[i]=buf[i]+(char)k;
		}
	}
return pbuf;
}
//-----------------------------------------------------------------------------------
long c24bitto32bit (char *Data)
{
long res=0;
res=(*Data)*256*256+(*(Data+1)*256)+(*(Data+2));
return res;
}

double cIEEE754toFloatC (char *Data)
{
BOOL sign;
char exp;
double res=0,zn=0.5, tmp;
unsigned char mask;
if (*(Data+0)&0x80) sign=TRUE; else sign=FALSE;
exp = ((*(Data+0)&0x7f)*2+(*(Data+1)&0x80)/0x80)-127;
for (int j=1;j<=3;j++)		// 23 значная мантисса
	{	
	 mask = 0x80;
	 for (int i=0;i<=7;i++)		
		{ 	 
		 if (j==1&&i==0) {res = res+1; mask = mask/2;}
		 else {
		 res = (*(Data+j)&mask)*zn/mask + res;
		 mask = mask/2; zn=zn/2; }
		}
	}
res = res * pow ((double)2,exp);
tmp = 1*pow ((double)10,-15);
if (res<tmp) res=0;
if (sign) res = -res;
return res;
}
//-----------------------------------------------------------------------------------
// 11111122 22223333 33444444 
// @ABCDEFGHIJKLMOPQRSTUVWXYZ[\]^ !"#$%&'()*+,-./0123456789:;<=>?
//-----------------------------------------------------------------------------------
char* Base64toString (char *Data,int lenght)
{
char buf[50];	
char *pbuf=buf;   
*pbuf = '\0';
int i=0,j=0;
for (i; i<lenght;i=i+3,j=j+4)
	{ 
	 buf[j]=(*(Data+i)&0xfc)/4;							
	 buf[j+1]=(*(Data+i)&0x3)*16 + (*(Data+i+1)&0xf0)/16; 
	 buf[j+2]=(*(Data+i+1)&0xf)*4 + (*(Data+i+2)&0xc0)/64;
	 buf[j+3]=(*(Data+i+2)&0x3f);
	 buf[j+4]='\0';
	}
for (i=0; i<j;i++)
	 if (buf[i]<32) buf[i]=64+buf[i];
return pbuf;
}

char* StringtoBase64 (char *Data,int lenght)
{
UL_INFO((LOGID, "lenght %i",lenght));
UL_INFO((LOGID, "input %s",Data));
char buf[50];	
char *pbuf=buf;
int i=0,j=0; 
for (i=0;i<24;i=i+3) {buf[i]=(char)0x82; buf[i+1]=(char)0x8; buf[i+2]=(char)0x20;}
for (i=0; i<lenght;i++)				// validate
	{
	 if ((*(Data+i))>63&&(*(Data+i))<95) (*(Data+i))=(*(Data+i))-64;
		else if ((*(Data+i))>31&&(*(Data+i))<64) (*(Data+i))=(*(Data+i));
			else	(*(Data+i))=32;
	}
for ( ;i<30;i++)	(*(Data+i))=32;
for (i=0; i<lenght;i=i+4,j=j+3)
	{ 	 
	 buf[j]=*(Data+i)*4 + (*(Data+i+1)&0x30)/16;
	 buf[j+1]=(*(Data+i+1)&0xf)*16+(*(Data+i+2)&0x3c)/4;
	 buf[j+2]=(*(Data+i+2)&0x3)*64+(*(Data+i+3)&0x3f);
	}
UL_INFO((LOGID, "output %s",buf));
buf[24]='\0';
return pbuf;
}
//-----------------------------------------------------------------------------------
char* ReadParam (char *SectionName,char *Value)
{
char buf[50]; 
char string1[50];
char *pbuf=buf;
unsigned int s_ok=0;
sprintf(string1,"[%s]",SectionName);
//UL_DEBUG((LOGID, "section name: %s",string1));
rewind (CfgFile);
	  while(!feof(CfgFile))
		 if(fgets(buf,50,CfgFile)!=NULL)
			if (strstr(buf,string1))
				{
				 s_ok=1; break;				 
				}
//else UL_ERROR((LOGID, ".ini file not found"));	 
if (s_ok)
	{
//	 UL_INFO((LOGID, "section %s found",SectionName));
	 while(!feof(CfgFile))
		{
		 if(fgets(buf,50,CfgFile)!=NULL&&strstr(buf,"[")==NULL&&strstr(buf,"]")==NULL)
			{
			 //UL_INFO((LOGID, "do %s / ishem %s",buf,Value));	// write in log
			 for (s_ok=0;s_ok<strlen(buf)-1;s_ok++)
				 if (buf[s_ok]==';') buf[s_ok+1]='\0';
			 if (strstr(buf,Value))
				{
				 for (s_ok=0;s_ok<strlen(buf)-1;s_ok++)
					if (s_ok>strlen(Value)) buf[s_ok-strlen(Value)-1]=buf[s_ok];
						 buf[s_ok-strlen(Value)-1]='\0';
			 	 //UL_INFO((LOGID, "value found = %s",buf));
				 //UL_INFO((LOGID, "Section name %s, value %s, che %s",SectionName,Value,buf));	// write in log
				 return pbuf;
				}
			}
		}	 	
	 if (SectionName=="Manufacturers"||SectionName=="Units") { buf[0]='u'; buf[1]='n'; buf[2]='k'; buf[3]='n'; buf[4]='o'; buf[5]='w'; buf[6]='n'; buf[7]='\0';	}
	 if (SectionName=="Models")	{ buf[0]='u'; buf[1]='n'; buf[2]='k'; buf[3]='n'; buf[4]='o'; buf[5]='w'; buf[6]='n'; buf[7]='\0';	}
 	 if (SectionName=="Port")	{ buf[0]='1'; buf[1]='\0';}
//	 UL_INFO((LOGID, "value = %s",buf));
	 return pbuf;
	}
else{
	 sprintf(buf, "error");			// if something go wrong return error
	 return pbuf;
	}	
}