// Machine generated IDispatch wrapper class(es) created by Microsoft Visual C++

// NOTE: Do not modify the contents of this file.  If this class is regenerated by
//  Microsoft Visual C++, your modifications will be overwritten.


#include "stdafx.h"
#include "explorer1.h"

/////////////////////////////////////////////////////////////////////////////
// CExplorer1

IMPLEMENT_DYNCREATE(CExplorer1, CWnd)

/////////////////////////////////////////////////////////////////////////////
// CExplorer1 properties

/////////////////////////////////////////////////////////////////////////////
// CExplorer1 operations
BEGIN_MESSAGE_MAP(CExplorer1, CWnd)
	ON_WM_RBUTTONDOWN()
	ON_WM_RBUTTONUP()
END_MESSAGE_MAP()

void CExplorer1::OnRButtonDown(UINT nFlags, CPoint point)
{
	// TODO: Add your message handler code here and/or call default
	MessageBox(_T("test"));

	CWnd::OnRButtonDown(nFlags, point);
}

void CExplorer1::OnRButtonUp(UINT nFlags, CPoint point)
{
	// TODO: Add your message handler code here and/or call default
	MessageBox(_T("up"));

	CWnd::OnRButtonUp(nFlags, point);
}

BOOL CExplorer1::PreTranslateMessage(MSG* pMsg)
{
	// TODO: Add your specialized code here and/or call the base class
	if ( WM_RBUTTONDOWN == pMsg->message || WM_RBUTTONUP == pMsg->message || 
		WM_KEYDOWN == pMsg->message )
	{
		return TRUE;
	}

	return CWnd::PreTranslateMessage(pMsg);
}
