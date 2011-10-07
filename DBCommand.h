// DBCommand.h: interface for the CDBCommand class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_DBCOMMAND_H__75286546_711A_4F4B_A158_DAE9A1FC9D3A__INCLUDED_)
#define AFX_DBCOMMAND_H__75286546_711A_4F4B_A158_DAE9A1FC9D3A__INCLUDED_

#include "DBConnect.h"

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class CDBCommand  
{
public:
	CDBCommand(_ConnectionPtr Conn = NULL);
	virtual ~CDBCommand();
	void Execute(const CString &strSQL);
	long GetIndex() const{
		return m_idx;
	}

private:
	_ConnectionPtr m_pConn;
	_CommandPtr m_pCmd;
	long m_idx;
};

#endif // !defined(AFX_DBCOMMAND_H__75286546_711A_4F4B_A158_DAE9A1FC9D3A__INCLUDED_)
