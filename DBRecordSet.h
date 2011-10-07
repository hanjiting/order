// DBRecordSet.h: interface for the CDBRecordSet class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_DBRECORDSET_H__6028A11E_8BA7_4FCC_B909_940A0D2DA3C5__INCLUDED_)
#define AFX_DBRECORDSET_H__6028A11E_8BA7_4FCC_B909_940A0D2DA3C5__INCLUDED_

#include "DBConnect.h"

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class CDBRecordSet  
{
public:
	CDBRecordSet(_ConnectionPtr Conn = NULL);
	virtual ~CDBRecordSet();
	BOOL Open(const CString &strSQL);
	void Close();
	BOOL Read();

	int GetValInt(int idx, int fmt = 0);
	double GetValDouble(int idx, int fmt = 0);
	bool GetValBool(int idx, int fmt = 0);
	CString GetValStr(int idx, int fmt = 0);
	COleDateTime GetValDateTime(int idx, int fmt = 0);

private:
	void dump_com_error(_com_error &e);
	int Conv2Int(_variant_t vtFld);
	double Conv2Double(_variant_t vtFld);
	bool Conv2Bool(_variant_t vtFld);
	void Conv2Str(_variant_t vtFld, CString & val, int fmt = 0);
	#define DT_DATETIME 0
	#define DT_DATEONLY 1
	#define DT_TIMEONLY	2
	COleDateTime Conv2DateTime(_variant_t vtFld);
	
private:
	_ConnectionPtr m_pConn;
	_RecordsetPtr m_pRs;
};

#endif // !defined(AFX_DBRECORDSET_H__6028A11E_8BA7_4FCC_B909_940A0D2DA3C5__INCLUDED_)
