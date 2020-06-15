#pragma once
#include <stdlib.h>

#ifdef _WIN32
#include <WinSock2.h>
#include <process.h>
#ifdef _UNICODE
#include <comdef.h>
#endif
typedef int socklen_t;
typedef void RET_TYPE;
#else
#include <errno.h>
#include <sys/types.h>
#include <sys/socket.h>
#include <netinet/in.h>
#include <arpa/inet.h>
#include <unistd.h>
#include <errno.h>
#include <string.h>
typedef unsigned int SOCKET;
#define INVALID_SOCKET  (SOCKET)(~0)
#define SOCKET_ERROR            (-1)
#ifndef _T
#define _T(x) x
#endif
typedef struct in_addr IN_ADDR;
typedef unsigned long       DWORD;
typedef int                 BOOL;
typedef unsigned char       BYTE;
typedef unsigned short      WORD;
typedef float               FLOAT;
typedef FLOAT               *PFLOAT;
typedef BOOL            *PBOOL;
typedef BOOL             *LPBOOL;
typedef BYTE            *PBYTE;
typedef BYTE             *LPBYTE;
typedef int             *PINT;
typedef int              *LPINT;
typedef WORD            *PWORD;
typedef WORD             *LPWORD;
typedef long             *LPLONG;
typedef DWORD           *PDWORD;
typedef DWORD            *LPDWORD;
typedef void             *LPVOID;

typedef int                 INT;
typedef unsigned int        UINT;
typedef unsigned int        *PUINT;
typedef const char* LPCTSTR, *LPCSTR;
typedef char* LPTSTR, *LPSTR;
typedef void* RET_TYPE;
inline int GetLastError()
{
	return errno;
}
#define closesocket(x) close(x)


#ifndef FALSE
#define FALSE               0
#endif

#ifndef TRUE
#define TRUE                1
#endif
#endif

//�������ڿ���̨����
class CPhdSocket
{
private:
	SOCKET m_hSocket;//���ĳ�Ա����

public:
	CPhdSocket();
	virtual ~CPhdSocket();

	operator SOCKET() const
	{
		return m_hSocket;
	}

	//�ر�socket
	void Close();

	// Summary:   �ö����Ƿ�Ϊ��
	// Time:	  2020��3��31�� peihaodong
	// Explain:	  
	bool IsNull() const;

	//����socket����
	//nSocketPort - �˿ڣ������0��������˿�
	//lpszSocketAddress - IP��ַ
	BOOL Create(UINT nSocketPort = 0, int nSocketType = SOCK_STREAM
		, LPCTSTR lpszSocketAddress = NULL);

	// Summary:   ����
	// Time:	  2020��3��30�� peihaodong
	// Explain:	  nConnectionBacklog - ����ͬʱ������5���ͻ��ˣ�����5����Ҫ�ȴ�
	BOOL Listen(int nConnectionBacklog = 5);

	// Summary:   ���գ������������пͻ������ӻ������µ�SocketPhd����
	// Time:	  2020��3��30�� peihaodong
	// Explain:	  ����SocketPhd�����IP�Ͷ˿�
	BOOL Accept(CPhdSocket& rConnectedSocket, LPTSTR szIP = NULL, 
		UINT *nPort = NULL);

	// Summary:   ���ӷ�������ͨ��������ǰ̨IP�Ͷ˿ڣ�
	// Time:	  2020��3��30�� peihaodong
	// Explain:	  
	BOOL Connect(LPCTSTR lpszHostAddress, UINT nHostPort);

	// Summary:   �������� ��õ�ַ�ṹ�� 
	// Time:	  2020��3��30�� peihaodong
	// Explain:	  
	int SendTo(const void* lpBuf, int nBufLen, UINT nHostPort,
		LPCTSTR lpszHostAddress = NULL);

	// Summary:   �������� �õ����Ͷ˵�IP�Ͷ˿�
	// Time:	  2020��3��30�� peihaodong
	// Explain:	  
	int ReceiveFrom(void* lpBuf, int nBufLen,
		LPTSTR rSocketAddress, UINT& rSocketPort);

	// Summary:   �������ݣ������ӵ�socket
	// Time:	  2020��3��30�� peihaodong
	// Explain:	  
	int Send(const void* lpBuf, int nBufLen, int nFlags = 0);

	// Summary:   �������ݣ������ӵ�socket
	// Time:	  2020��3��30�� peihaodong
	// Explain:	  
	int Receive(void* lpBuf, int nBufLen, int nFlags = 0);
	
	// Summary:   �õ�socket��Ϣ
	// Time:	  2020��3��30�� peihaodong
	// Explain:	  
	BOOL GetPeerName(LPTSTR rSocketAddress, UINT& rSocketPort);
	BOOL GetSockName(LPTSTR rSocketAddress, UINT& rSocketPort);

};

