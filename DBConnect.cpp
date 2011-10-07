// DBConnect.cpp: implementation of the CDBConnect class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "DBConnect.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CDBConnect *CDBConnect::m_Instance = NULL;

CDBConnect::CDBConnect()
{
	m_pConn = NULL;

	char m_pBuf[MAX_PATH];
	GetCurrentDirectory(MAX_PATH,m_pBuf);
	strcat(m_pBuf,"\\MgrTbl.mdb");
	m_DbPath = m_pBuf;
	
	m_SizeBeforeCompress = 0;
	m_SizeAfterCompress = 0;
}

_ConnectionPtr CDBConnect::GetActiveConnection()
{
	if(NULL == m_pConn)
	{
		HRESULT hr;
		CString strTmp, strTmp1;
		try
		{
			// 创建Connection对象
			// hr = m_pConn.CreateInstance("ADODB.Connection");
			hr = m_pConn.CreateInstance(__uuidof(Connection));
			if(SUCCEEDED(hr))
			{
				// 连接数据库
				// 针对ACCESS2000环境，Provider=Microsoft.Jet.OLEDB.4.0;
				// 针对ACCESS97，      Provider=Microsoft.Jet.OLEDB.3.51; 
				strTmp = _T("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=");
				strTmp1 = strTmp + m_DbPath;
				//strTmp = _T(";Jet OLEDB:Database Password=gongzuoshi");
				//strTmp1 = strTmp1 + strTmp;
				hr = m_pConn->Open((LPSTR)(LPCTSTR)strTmp1,"","",adModeUnknown);
				//hr = m_pConn->Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=MgrTbl.mdb","","",adModeUnknown);
			}
		}
		catch(_com_error &e)///捕捉异常
		{
			CString errormessage;
			errormessage.Format("连接数据库失败!\r\n错误信息:%s",e.ErrorMessage());
			AfxMessageBox(errormessage);///显示错误信息
		}
	}

	return m_pConn;
}

CDBConnect::~CDBConnect()
{
}

bool CDBConnect::SetDbFile(CString const &strFile)
{
	if(m_DbPath == strFile)
	{
		return false;
	}

	m_DbPath = strFile;

	return true;
}

CDBConnect *CDBConnect::GetInstance()
{
	if(NULL == m_Instance)
	{
		m_Instance = new CDBConnect;
	}

	return m_Instance;
}


void CDBConnect::ExitInstance()
{
	if(NULL != m_Instance)
	{
		m_Instance->CloseConnection();
		delete m_Instance;
		m_Instance = NULL;
	}
	::CoUninitialize();
}

void CDBConnect::InitInstance()
{
	::CoInitialize(NULL);
}

void CDBConnect::CloseConnection()
{
	if(NULL != m_pConn)
	{
		if(m_pConn->GetState() != adStateClosed)
		{
			m_pConn->Close();
		}
		m_pConn.Release();
		m_pConn = NULL;
		
	}
}

bool CDBConnect::CompressDB(CString *strFile)
{
	CString strTmp, strTmp1;
	
	if(NULL == strFile)
	{
		CloseConnection();
		strTmp1 = m_DbPath;
	}
	else
	{
		strTmp1 = *strFile;
	}

	CFile file;
	//get file size
	file.Open(strTmp1, CFile::modeRead);
	m_SizeBeforeCompress = file.GetLength();
	file.Close();

	//compress
	strTmp = _T("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=");
	strTmp = strTmp + strTmp1;
	//strTmp = strTmp + _T(";Jet OLEDB:Database Password=gongzuoshi");
	try{
		IJetEnginePtr jet(__uuidof(JetEngine));	
		jet->CompactDatabase(
			(LPSTR)(LPCTSTR)strTmp, 
			"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\TmpDb20110902.mdb;"\
			"Jet OLEDB:Engine Type=5");	//;Jet OLEDB:Database Password=gongzuoshi
	}
	catch(_com_error &e)
	{
#ifdef _DEBUG
		CString errormessage;
		errormessage.Format("压缩失败:%s",e.ErrorMessage());
		AfxMessageBox(errormessage);
		return false;
#endif
	}

	::CopyFile("C:\\TmpDb20110902.mdb", strTmp1, false);

	::DeleteFile("C:\\TmpDb20110902.mdb");

	//get file size
	file.Open(strTmp1, CFile::modeRead);
	m_SizeAfterCompress = file.GetLength();
	file.Close();

	return true;
}
