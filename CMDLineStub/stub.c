#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <windows.h>

int APIENTRY WinMain(HINSTANCE inst, HINSTANCE prvinst, LPSTR cmdline, int cmdshow){
	LPSTR _fullcmdline;
	static char fullcmdline[32768];
	int pos;
	HANDLE hstdin;
	HANDLE hstdout;
	HANDLE hstderr;
	STARTUPINFOA startinfo;
	PROCESS_INFORMATION procinfo;
	BOOL success;
	DWORD exitcode;

	// Append .exe to argv[0] in commandline
	_fullcmdline = GetCommandLineA();
	memset (fullcmdline, 0, 32768);
	strncpy (fullcmdline, _fullcmdline, 32767);
	fullcmdline[32767] = 0;
	pos = strlen(fullcmdline) - strlen(cmdline);

	while (pos > 0 && fullcmdline[pos - 1] == ' '){
		pos--;
	}

	if (memcmp (fullcmdline + pos - 4, ".com", 4) == 0){
		memcpy (fullcmdline + pos - 4, ".exe", 4);
	} else if (memcmp (fullcmdline + pos - 5, ".com\"", 5) == 0){
		memcpy (fullcmdline + pos - 5, ".exe\"", 5);
	} else if (memcmp (fullcmdline + pos - 5, ".exe\"", 5) == 0 ||
			   memcmp (fullcmdline + pos - 4, ".exe", 4) == 0){
		printf ("Oh dear.  I should end with .com\n");
		ExitProcess(1);
	} else if (memcmp (fullcmdline + pos - 1, "\"", 1) == 0){
		memcpy (fullcmdline + pos - 1, ".exe\"", 5);
		pos += 4;
	} else {
		memcpy (fullcmdline + pos, ".exe", 4);
		pos += 4;
	}

	fullcmdline[pos++] = ' ';
	memset (fullcmdline + pos, 0, 32768 - pos);
	strncpy (fullcmdline + pos, cmdline, 32768 - pos);
    
	// Initialize STARTUPINFO
	memset (&startinfo, 0, sizeof(STARTUPINFOA));
	startinfo.cb = sizeof(STARTUPINFOA);
	GetStartupInfoA (&startinfo);
	startinfo.cb = sizeof(STARTUPINFOA);
	startinfo.lpReserved = NULL;
	startinfo.lpTitle = NULL;
	startinfo.dwX = 0;
	startinfo.dwY = 0;
	startinfo.dwXSize = 0;
	startinfo.dwYSize = 0;
	startinfo.dwXCountChars = 0;
	startinfo.dwYCountChars = 0;
	startinfo.dwFillAttribute = 0;
	startinfo.dwFlags = 0;
	startinfo.wShowWindow = 0;
	startinfo.cbReserved2 = 0;
	startinfo.lpReserved2 = NULL;
	startinfo.hStdInput = (HANDLE)0;
	startinfo.hStdOutput = (HANDLE)0;
	startinfo.hStdError = (HANDLE)0;

	memset (&procinfo, 0, sizeof(PROCESS_INFORMATION));

	// Execute the app
	success = CreateProcess (
		NULL,
		fullcmdline,
		NULL,
		NULL,
		TRUE,
		DETACHED_PROCESS,
		NULL,
		NULL,
		&startinfo,
		&procinfo);

	if (success == 0){
		printf ("Oh dear. Couldn't execute the program: %08X\n", GetLastError());
		ExitProcess(1);
	}

	exitcode = WaitForSingleObject (procinfo.hProcess, INFINITE);
	if (exitcode != WAIT_OBJECT_0){
		printf ("Oh dear. Waiting for the process returned %d\n", exitcode);
		ExitProcess(exitcode);
	}

	success = GetExitCodeProcess (procinfo.hProcess, &exitcode);
	if (success == 0){
		printf ("Oh dear. Couldn't get the exit code: %08X\n", GetLastError());
		ExitProcess(1);
	}

	ExitProcess(exitcode);
}

