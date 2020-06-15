#include "StdAfx.h"
#include "CPhdSocket.h"
#include <ws2tcpip.h>
#ifdef _WIN32
#pragma comment (lib,"ws2_32.lib")
#endif

CPhdSocket::CPhdSocket()
{
#ifdef _WIN32
	WSAData wd;
	WSAStartup(0x0202, &wd);
#endif
	m_hSocket = INVALID_SOCKET;
}

CPhdSocket::~CPhdSocket()
{
	Close();
}

void CPhdSocket::Close()
{
	closesocket(m_hSocket);
	m_hSocket = INVALID_SOCKET;
}

bool CPhdSocket::IsNull() const
{
	if (m_hSocket == INVALID_SOCKET)
		return true;
	else
		return false;
}

BOOL CPhdSocket::Create(UINT nSocketPort /*= 0*/,
	int nSocketType /*= SOCK_STREAM */, 
	LPCTSTR lpszSocketAddress /*= NULL*/)
{
	//����socket
	m_hSocket = socket(AF_INET, nSocketType, 0);
	if (m_hSocket == INVALID_SOCKET)
	{
		return FALSE;
	}
	//������ַ��Ϣ�ṹ�壺IP�Ͷ˿�
	sockaddr_in sa = { AF_INET,htons(nSocketPort) };
	if (lpszSocketAddress)
	{
		//IP��ַ���뵽��ַ�ṹ���У�InetPton���ڿ��ֽڣ�inet_pton���ڶ��ֽڣ�
		InetPton(AF_INET, lpszSocketAddress, &sa.sin_addr.s_addr);
	}
	//socket�󶨵�ַ�ṹ��
	return !bind(m_hSocket, (sockaddr*)&sa, sizeof(sa));
}

BOOL CPhdSocket::Listen(int nConnectionBacklog /*= 5*/)
{
	return !listen(m_hSocket, nConnectionBacklog);
}

BOOL CPhdSocket::Accept(CPhdSocket& rConnectedSocket,
	LPTSTR szIP /*= NULL*/, UINT *nPort /*= NULL*/)
{
	//���������ַ�ṹ��
	sockaddr_in sa = { AF_INET };
	socklen_t nLen = sizeof(sa);
	rConnectedSocket.m_hSocket = accept(m_hSocket, (sockaddr*)&sa, &nLen);
	if (INVALID_SOCKET == rConnectedSocket.m_hSocket)
		return FALSE;
	if (szIP)
	{
		//ͨ���õ�ַ�ṹ��õ�IP
		wchar_t strIP[100] = { 0 };
		inet_ntop(AF_INET, (void *)&sa.sin_addr, _bstr_t(strIP), sizeof(strIP));
	}
	if (nPort)
		*nPort = htons(sa.sin_port);
	return TRUE;
}

BOOL CPhdSocket::Connect(LPCTSTR lpszHostAddress, UINT nHostPort)
{
	sockaddr_in sa = { AF_INET,htons(nHostPort) };
	//IP��ַ���뵽��ַ�ṹ���У�InetPton���ڿ��ֽڣ�inet_pton���ڶ��ֽڣ�
	InetPton(AF_INET, lpszHostAddress, &sa.sin_addr.s_addr);
	return !connect(m_hSocket, (sockaddr*)&sa, sizeof(sa));
}

int CPhdSocket::SendTo(const void* lpBuf, int nBufLen, UINT nHostPort,
	LPCTSTR lpszHostAddress /*= NULL*/)
{
	sockaddr_in sa = { AF_INET,htons(nHostPort) };
	//IP��ַ���뵽��ַ�ṹ���У�InetPton���ڿ��ֽڣ�inet_pton���ڶ��ֽڣ�
	InetPton(AF_INET, lpszHostAddress, &sa.sin_addr.s_addr);
	sendto(m_hSocket, (const char*)lpBuf, nBufLen, 0, (sockaddr*)&sa,
		sizeof(sa));
	return 0;
}

int CPhdSocket::ReceiveFrom(void* lpBuf, int nBufLen, 
	LPTSTR rSocketAddress, UINT& rSocketPort)
{
	sockaddr_in sa = { AF_INET };
	socklen_t nLen = sizeof(sa);
	int nRet = recvfrom(m_hSocket, (char*)lpBuf, nBufLen, 0, 
		(sockaddr*)&sa, &nLen);
	if (nRet <= 0)
		return nRet;

	rSocketPort = htons(sa.sin_port);
	if (rSocketAddress)
	{
		//ͨ���õ�ַ�ṹ��õ�IP
		wchar_t strIP[100] = { 0 };
		inet_ntop(AF_INET, (void *)&sa.sin_addr,
			_bstr_t(strIP), sizeof(strIP));
	}
	return nRet;
}

int CPhdSocket::Send(const void* lpBuf, int nBufLen, int nFlags /*= 0*/)
{
	return send(m_hSocket, (LPCSTR)lpBuf, nBufLen, nFlags);
}

int CPhdSocket::Receive(void* lpBuf, int nBufLen, int nFlags /*= 0*/)
{
	return recv(m_hSocket, (LPSTR)lpBuf, nBufLen, nFlags);
}

BOOL CPhdSocket::GetPeerName(LPTSTR rSocketAddress, UINT& rSocketPort)
{
	sockaddr_in sa = { AF_INET };
	socklen_t nLen = sizeof(sa);
	if (getpeername(m_hSocket, (sockaddr*)&sa, &nLen) < 0)
		return FALSE;
	rSocketPort = htons(sa.sin_port);
	if (rSocketAddress)
	{
		//ͨ���õ�ַ�ṹ��õ�IP
		wchar_t strIP[100] = { 0 };
		inet_ntop(AF_INET, (void *)&sa.sin_addr,
			_bstr_t(strIP), sizeof(strIP));
	}
	return TRUE;
}

BOOL CPhdSocket::GetSockName(LPTSTR rSocketAddress, UINT& rSocketPort)
{
	sockaddr_in sa = { AF_INET };
	socklen_t nLen = sizeof(sa);
	if (getsockname(m_hSocket, (sockaddr*)&sa, &nLen) < 0)
		return FALSE;
	rSocketPort = htons(sa.sin_port);
	if (rSocketAddress)
	{
		//ͨ���õ�ַ�ṹ��õ�IP
		wchar_t strIP[100] = { 0 };
		inet_ntop(AF_INET, (void *)&sa.sin_addr,
			_bstr_t(strIP), sizeof(strIP));
	}
	return TRUE;
}
