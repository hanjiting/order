// DBCommand.cpp: implementation of the CDBCommand class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "DBCommand.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CDBCommand::CDBCommand(_ConnectionPtr Conn)
{
	m_pConn = Conn;
	m_pCmd = NULL;

	if(NULL == Conn)
	{
		m_pConn = CDBConnect::GetInstance()->GetActiveConnection();
	}
	
	m_pCmd.CreateInstance(__uuidof(Command));

	m_idx = -1;
}

CDBCommand::~CDBCommand()
{
	if(NULL != m_pCmd)
	{
		m_pCmd.Release();
	}

	m_pCmd = NULL;
}

void CDBCommand::Execute(const CString &strSQL)
{
	ASSERT(!strSQL.IsEmpty());
	ASSERT(m_pConn != NULL);
	ASSERT(m_pConn->GetState() != adStateClosed);

	m_pCmd->ActiveConnection = m_pConn;
	m_pCmd->CommandText = _bstr_t(strSQL);
	m_pCmd->CommandType = adCmdStoredProc;

	m_pCmd->Execute(NULL, NULL, adCmdText);

	if(strSQL.Mid(0, strlen("Insert")).CompareNoCase("insert") == 0)
	{
		CString strSQL1;
		strSQL1.Format("select @@Identity as idx");
		m_pCmd->CommandText = _bstr_t(strSQL1);
		_RecordsetPtr pRs = m_pCmd->Execute(NULL, NULL, adCmdText);
		_variant_t vtFld;
		vtFld = pRs->Fields->GetItem((long)0)->Value;
		m_idx = vtFld.ulVal;
	}
}