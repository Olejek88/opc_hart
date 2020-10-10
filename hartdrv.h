#include <opcda.h>
#define HARTCOM_DATALEN_MAX 25
#define HARTCOM_ID_MAX	4

typedef struct _HartCom HartCom;
typedef struct _Hart Hart;

struct _HartCom {
  char *name;
  unsigned int getCmd;
  unsigned int setCmd;
  VARTYPE dtype;
  unsigned int start;
  unsigned int num;
  BOOL scan_it;
};

HartCom HartCommU[] =
 {{"Manufacturer Code",	0,	-1,	VT_UI1,	1,	1,	true},
  {"Device Code",		0,	-1,	VT_UI1,	2,	1,	true},
  {"Commands rev.",		0,	-1,	VT_I2,	4,	1,	true},
  {"Device com. rev.",	0,	-1,	VT_I2,	5,	1,	true},
  {"Soft rev.",			0,	-1,	VT_I2,	6,	1,	true},
  {"Hard rev.",			0,	-1,	VT_I2,	7,	1,	true},
  {"Serial number",		0,	-1,	VT_I4,	9,	3,	true},
  {"PV unit code",		1,	-1,	VT_UI1,	0,	1,	true},
  {"PV",				1,	-1,	VT_R4,	1,	4,	true},
  {"SV unit code",		3,	-1,	VT_UI1,	9,	1,	true},
  {"SV",				3,	-1,	VT_R4,	10,	4,	true},
  {"TV unit code",		3,	-1,	VT_UI1,	14,	1,	true},
  {"TV",				3,	-1,	VT_R4,	15,	4,	true},
  {"QV unit code",		3,	-1,	VT_UI1,	19,	1,	true},
  {"QV",				3,	-1,	VT_R4,	20,	4,	true},
  {"PV%Range",			2,	-1,	VT_R4,	4,	4,	true},
  {"LoopCurrent",		2,	40,	VT_R4,	0,	4,	true},
  {"Message",			12,	17,	VT_BSTR,0,	25,	true},
  {"Shortform adr",		6,	6,	VT_UI1, 0,	1,	true},
  {"PV Damping",		15,	34,	VT_R4,	11,	4,	true},
  {"Upper sensor lim",	14,	-1,	VT_R4,	4,	4,	true},
  {"Lower sensor lim",	14,	-1,	VT_R4,	8,	4,	true},
  {"PV minimum span",	14,	-1,	VT_R4,	12,	4,	true},
  {"Descriptor",		13,	18,	VT_BSTR,6,	12,	true},
  {"TAG",				13,	18,	VT_BSTR,0,	6,	true},
  {"Date",				13,	18,	VT_BSTR,18,	3,	true},
  {"Lower range value",	15,	-1,	VT_R4,	7,	4,	true},
  {"Upper range value",	15,	-1,	VT_R4,	3,	4,	true},

  {"Manufacturer Name",	-1,	-1,	VT_BSTR,1,	1,	false},
  {"Device Name",	-1,	-1,	VT_BSTR,2,	1,	false},
  {"PV unit",		-1,	-1,	VT_BSTR,0,	1,	false},
  {"SV unit",		-1,	-1,	VT_BSTR,9,	1,	false},
  {"TV unit",		-1,	-1,	VT_BSTR,14,	1,	false},
  {"QV unit",		-1,	-1,	VT_BSTR,19,	1,	false}};

struct _Hart {
  int cv_size;          // size (quantity scanned commands) 
  LPSTR *cv;			// pointer to value
  int *cv_cmdid;        // command identitificator
  BOOL *cv_status;		// status
  int ids[HARTCOM_ID_MAX+1];
  int idnum;

_Hart()//: mtime(0), idnum(0)
  {
    int i;
    
    for (i = 1, cv_size = 0; HartCommU[i].name != NULL; i++)
	cv_size += 1;						// scanned command + 1

    cv = new LPSTR[cv_size];			// ???
    cv_status = new BOOL[cv_size];		// massive status
    cv_cmdid = new int[cv_size];        // massive id

    int cv_i;				
    for (i = 0, cv_i = 0; HartCommU[i].name != NULL; i++) {
	cv[cv_i] = new char[HARTCOM_DATALEN_MAX+1];	// init tag	
	cv_status[cv_i] = FALSE;					// init status
	cv_cmdid[cv_i] = i;							// init id
	cv_i++;
    }
  }

~_Hart()						// destructor
  {
    int i;
    for (i =0; i < cv_size; i++)
      delete[] cv[i];           // free memory

    delete[] cv_status;
    delete[] cv_cmdid;
    delete[] cv;
  }
};
