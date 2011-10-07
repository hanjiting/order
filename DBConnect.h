// DBConnect.h: interface for the CDBConnect class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_DBCONNECT_H__0A762733_8D5D_4CA2_A686_B8C136109C25__INCLUDED_)
#define AFX_DBCONNECT_H__0A762733_8D5D_4CA2_A686_B8C136109C25__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#pragma warning (disable:4146)
#import "c:\program files\common files\system\ado\msado15.dll" no_namespace rename("EOF","adoEOF")
#import "C:\program files\common files\system\ado\msjro.dll"  no_namespace rename("ReplicaTypeEnum", "_ReplicaTypeEnum")
#pragma warning (default:4146)

class CDBConnect  
{
public:
	bool CompressDB(CString *strFile = NULL);
	static void InitInstance(void);
	static CDBConnect *GetInstance();
	static void ExitInstance();
	virtual ~CDBConnect();
	_ConnectionPtr GetActiveConnection();
	void CloseConnection(void);
	bool SetDbFile(CString const &strFile);
	DWORD GetSizeBeforeCompress(){
		return m_SizeBeforeCompress;
	}
	DWORD GetSizeAfterCompress(){
		return m_SizeAfterCompress;
	}

private:
	CDBConnect();
private:
	static CDBConnect *m_Instance;
	_ConnectionPtr m_pConn;
	CString m_DbPath;
	DWORD m_SizeBeforeCompress;
	DWORD m_SizeAfterCompress;
};

#endif // !defined(AFX_DBCONNECT_H__0A762733_8D5D_4CA2_A686_B8C136109C25__INCLUDED_)
