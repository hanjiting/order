// DBRecordSet.cpp: implementation of the CDBRecordSet class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "DBRecordSet.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

#define CNVBUF_SIZE 20

CDBRecordSet::CDBRecordSet(_ConnectionPtr Conn)
{
	m_pConn = Conn;
	m_pRs = NULL;

	if(NULL == Conn)
	{
		m_pConn = CDBConnect::GetInstance()->GetActiveConnection();
	}

	m_pRs.CreateInstance(__uuidof(Recordset));
}

CDBRecordSet::~CDBRecordSet()
{
	if(NULL != m_pRs)
	{
		Close();
		m_pRs.Release();
		m_pRs = NULL;
	}
}

////////////////////////////////////////////////////////////////////////////////////////////////////////
//
// HRESULT Recordset15::Open ( const _variant_t & Source, 
//							   const _variant_t & ActiveConnection, 
//							   enum CursorTypeEnum CursorType, 
//							   enum LockTypeEnum LockType, 
//							   long Options )
//
// CursorType
//   - adOpenForwardOnly   0   ȱʡֵ������һ��ֻ����ǰ�ƶ����α꣨Forward Only����
//                             ���ֹ��ֻ����ǰ�����¼����������MoveNext��ǰ����,
//                             ���ַ�ʽ�����������ٶȡ�������BookMark,RecordCount,
//                             AbsolutePosition,AbsolutePage������ʹ��
//							   (ֱ���ڻ����ϲ���)
//   - adOpenKeyset        1   ����һ��Keyset���͵��αꡣ�������ֹ��ļ�¼��������
//							   �����û���������ɾ�������������ڸ���ԭ�м�¼�Ĳ����ǿɼ��ġ�
//							   (��Tempdb��Ϊ�α��¼����һ����Ӧ��Keyset)
//   - adOpenDynamic       2   ����һ��Dynamic���͵��αꡣ�������ݿ�Ĳ������������ڸ��û���¼���Ϸ�Ӧ������
//							   (ֱ���ڻ����ϲ���)
//   - adOpenStatic        3   ����һ��Static���͵��αꡣ��Ϊ��ļ�¼������һ����̬���ݣ�
//                             �������û���������ɾ�������²�������ļ�¼����˵�ǲ��ɼ��ġ�
//							   (��Tempdb��Ϊ�α��¼����һ������)
//
// LockType:
//	 - adLockReadOnly     1   ȱʡֵ��Recordset������ֻ����ʽ�������޷�����AddNew��Update��Delete�ȷ���
//   - adLockPrssimistic  2   �����ڸ���ʱ��ϵͳ����ס�����û��Ķ������Ա�������һ���ԣ������ȫ����������
//   - adLockOptimistic   3   �����ڸ���ʱ��ϵͳ��������ס�����û��Ķ�����ֻ���������Update����ʱ��������¼��
//                            �ڴ�֮ǰ�������û����Զ����ݽ�������ɾ���ĵĲ�����
//   - adLockBatchOptimistic 4 �༭ʱ��¼���������������û����뽫CursorLocation����
//                             ��ΪadUdeClientBatch���ܶ����ݽ�������ɾ���ĵĲ�����
////////////////////////////////////////////////////////////////////////////////////////////////////////

BOOL CDBRecordSet::Open(const CString &strSQL)
{
	ASSERT(!strSQL.IsEmpty());
	ASSERT(m_pConn != NULL);
	ASSERT(m_pConn->GetState() != adStateClosed);

	Close();

#ifdef _DEBUG
	if(strSQL.Mid(0, strlen("Select")).CompareNoCase("select") != 0)
	{
		AfxMessageBox("ֻ����ͨ��CDBRecordSetִ��select�����ȡ�ü�¼��");
		return FALSE;
	}
#endif
	m_pRs->Open((LPCSTR)strSQL, _variant_t((IDispatch*)m_pConn, TRUE), 
						adOpenDynamic, adLockReadOnly, adCmdUnknown);

	return (!m_pRs->GetadoEOF());
}

void CDBRecordSet::Close()
{
	if(m_pRs->GetState() != adStateClosed)
	{
		m_pRs->Close();
	}
}

BOOL CDBRecordSet::Read()
{
	if(m_pRs->GetadoEOF())
	{
		return FALSE;
	}

	m_pRs->MoveNext();

	if(m_pRs->GetadoEOF())
	{
		return FALSE;
	}

	return TRUE;
}

int CDBRecordSet::GetValInt(int idx, int fmt)
{
	int val = 0;
	_variant_t vtFld;
	_variant_t vtIndex;

	vtIndex.vt = VT_I2;
	vtIndex.iVal = idx;
	
	try
	{
		vtFld = m_pRs->Fields->GetItem(vtIndex)->Value;
		val = Conv2Int(vtFld);
	}
	catch(_com_error &e)
	{
		dump_com_error(e);
	}

	return val;
}

double CDBRecordSet::GetValDouble(int idx, int fmt)
{
	double val = 0.0;
	_variant_t vtFld;
	_variant_t vtIndex;

	vtIndex.vt = VT_I2;
	vtIndex.iVal = idx;
	
	try
	{
		vtFld = m_pRs->Fields->GetItem(vtIndex)->Value;
		val = Conv2Double(vtFld);
	}
	catch(_com_error &e)
	{
		dump_com_error(e);
	}

	return val;
}

bool CDBRecordSet::GetValBool(int idx, int fmt)
{
	bool val = false;
	_variant_t vtFld;
	_variant_t vtIndex;
	
	vtIndex.vt = VT_I2;
	vtIndex.iVal = idx;
	
	try
	{
		vtFld = m_pRs->Fields->GetItem(vtIndex)->Value;
		val = Conv2Bool(vtFld);
	}
	catch(_com_error &e)
	{
		dump_com_error(e);
	}

	return val;
}

CString CDBRecordSet::GetValStr(int idx, int fmt)
{
	CString val;

	_variant_t vtFld;
	_variant_t vtIndex;
	
	vtIndex.vt = VT_I2;
	vtIndex.iVal = idx;
	
	try
	{
		vtFld = m_pRs->Fields->GetItem(vtIndex)->Value;
		Conv2Str(vtFld, val, fmt);
	}
	catch(_com_error &e)
	{
		dump_com_error(e);
	}

	return val;
}

COleDateTime CDBRecordSet::GetValDateTime(int idx, int fmt)
{
	COleDateTime val;

	_variant_t vtFld;
	_variant_t vtIndex;
	
	vtIndex.vt = VT_I2;
	vtIndex.iVal = idx;
	
	try
	{
		vtFld = m_pRs->Fields->GetItem(vtIndex)->Value;
		val = Conv2DateTime(vtFld);
	}
	catch(_com_error &e)
	{
		dump_com_error(e);
	}

	return val;
}

////////////////////////////////////////////////////////////////////////////////////
//
//
//
////////////////////////////////////////////////////////////////////////////////////
void CDBRecordSet::dump_com_error(_com_error &e)
{
#ifdef _DEBUG
	AfxMessageBox( e.Description(), MB_OK | MB_ICONERROR );
#endif
}

int CDBRecordSet::Conv2Int(_variant_t vtFld)
{
	int val;

	switch(vtFld.vt)
	{
	case VT_INT:
		val = vtFld.intVal;
		break;
	case VT_I4:		//long
		val = vtFld.lVal;
		break;
	case VT_NULL:
	case VT_EMPTY:
		val = 0;
		break;
	case VT_R8:		//double
		val = (int)vtFld.dblVal;
		break;
	case VT_BSTR:	//BSTR
		{
			char ConvBuf[CNVBUF_SIZE];
			WideCharToMultiByte(CP_ACP, 0, vtFld.bstrVal, -1, ConvBuf, sizeof(ConvBuf), NULL, NULL);
			val = atoi(ConvBuf);
		}
		break;
	case VT_BOOL:	//VARIANT_BOOL
		val = vtFld.boolVal;
		break;
	case VT_UI1:	//unsigned char
		val = vtFld.bVal;
		break;
	case VT_I2:		//short
		val = vtFld.iVal;
		break;
	case VT_R4:		//float
		val = (int)vtFld.fltVal;
		break;
	default :
		val = vtFld.iVal;
		break;
	}

	return val;
}

double CDBRecordSet::Conv2Double(_variant_t vtFld)
{
	double val = 0.0;

	switch(vtFld.vt)
	{
	case VT_R8:
		val = vtFld.dblVal;
		break;
	case VT_NULL:
	case VT_EMPTY:
		val = 0.0;
		break;
	case VT_BSTR:	//BSTR
		{
			char ConvBuf[CNVBUF_SIZE];
			WideCharToMultiByte(CP_ACP, 0, vtFld.bstrVal, -1, ConvBuf, sizeof(ConvBuf), NULL, NULL);
			val = atof(ConvBuf);
		}
		break;
	case VT_I4:		//long
		val = vtFld.lVal;
		break;
	case VT_UI1:	//unsigned char
		val = vtFld.bVal;
		break;
	case VT_INT:
		val = vtFld.intVal;
		break;
	case VT_I2:		//short
		val = vtFld.iVal;
		break;
	case VT_R4:
		val = vtFld.fltVal;
		break;
// 	case VT_DECIMAL:
// 		val = vtFld.decVal.Lo32;
// 		val *= (vtFld.decVal.sign == 128)? -1 : 1;
// 		val /= pow(10, vtFld.decVal.scale); 
// 		break;
	case VT_CY:
		vtFld.ChangeType(VT_R8);
		val = vtFld.dblVal;
		break;
	default:
		val = vtFld.dblVal;
	}

	return val;
}

bool CDBRecordSet::Conv2Bool(_variant_t vtFld)
{
	bool val = false;

	switch(vtFld.vt) 
	{
	case VT_BOOL:
		val = vtFld.boolVal == VARIANT_TRUE? true: false;
		break;
	case VT_EMPTY:
	case VT_NULL:
		val = false;
		break;
	case VT_BSTR:	//BSTR
		{
			CString strtmp;
			char ConvBuf[CNVBUF_SIZE];
			WideCharToMultiByte(CP_ACP, 0, vtFld.bstrVal, -1, ConvBuf, sizeof(ConvBuf), NULL, NULL);
			strtmp = ConvBuf;
			if(!strtmp.CompareNoCase("true"))	//is true
			{
				val = true;
			}
		}
		break;
	default:
		//val = false;
		if(0 != vtFld.iVal)
		{
			val = true;
		}
		break;
	}

	return val;
}

void CDBRecordSet::Conv2Str(_variant_t vtFld, CString & val, int fmt)
{
	//char ConvBuf[CNVBUF_SIZE];
	double dbval = 0.0;
	
	switch(vtFld.vt) 
	{
	////////////////////////////return...
	case VT_BSTR:
		val = vtFld.bstrVal;
		return;	//return
	case VT_DECIMAL:
		val.Format("%d", vtFld.decVal);
		return;	//return
	case VT_DATE:
		{
			COleDateTime dt(vtFld);
			switch(fmt)
			{
			default :
				val = dt.Format("%Y-%m-%d %H:%M:%S");
				break;
			case DT_DATEONLY:
				val = dt.Format("%Y-%m-%d");
				break;
			case DT_TIMEONLY:
				val = dt.Format("%H:%M:%S");
				break;
			}
		}
		return;	//return
	case VT_CY:
		{
			COleCurrency cy(vtFld);
			val = cy.Format();
		}
		return;	//return
	case VT_BOOL:
		val = (vtFld.boolVal == VARIANT_TRUE ? "true" : "false");
		return;	//return
	case VT_EMPTY:
	case VT_NULL:
	default:
		val.Empty();
		return;	//return
	////////////////////////////break...
	case VT_R4:
		dbval = vtFld.fltVal;
		//val.Format("%f",vtFld.fltVal);
		break;
	case VT_R8:
		dbval = vtFld.dblVal;
		//val.Format("%f", vtFld.dblVal);
		break;
	case VT_I2:
	case VT_UI1:
		dbval = vtFld.iVal;
		//val.Format("%d", vtFld.iVal);
		break;
	case VT_INT:
		dbval = vtFld.intVal;
		//val.Format("%d", vtFld.intVal);
		break;
	case VT_I4:
		dbval = vtFld.lVal;
		//val.Format("%d", vtFld.lVal);
		break;
	case VT_UI4:
		dbval = vtFld.ulVal;
		//val.Format("%u", vtFld.ulVal);
		break;
	}
	CString strtmp;
	strtmp.Format("%%.%df", fmt);
	val.Format(strtmp, dbval);
}

COleDateTime CDBRecordSet::Conv2DateTime(_variant_t vtFld)
{
	COleDateTime val;
	
	switch(vtFld.vt) 
	{
	case VT_DATE:
		{
			COleDateTime dt(vtFld);
			val = dt;
		}
		break;
	case VT_BSTR:	//BSTR
		{
			char ConvBuf[CNVBUF_SIZE];
			WideCharToMultiByte(CP_ACP, 0, vtFld.bstrVal, -1, ConvBuf, sizeof(ConvBuf), NULL, NULL);
			val.ParseDateTime(ConvBuf);
		}
		break;
	case VT_EMPTY:
	case VT_NULL:
	default:
		val.SetStatus(COleDateTime::null);
		break;
	}
	
	return val;
}

