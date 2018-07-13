#define WIN32_LEAN_AND_MEAN
#include <Windows.h>
#include <shellapi.h>
#include <objbase.h>
#include <string>

#pragma comment(linker, "/SUBSYSTEM:windows /ENTRY:mainCRTStartup")

int main(int argc, const char *argv[])
{
    CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);

    std::string args = "/x";

    for (int i = 1; i < argc; i++) {
        args += " \"" + std::string(argv[i]) + "\"";
    }

    //std::string directory = "C:\\Program Files (x86)\\Microsoft Office\\root\\Office16";  // Excel 2016
    std::string directory = "C:\\Program Files (x86)\\Microsoft Office\\Office14";          // Excel 2010
    std::string executable = directory + "\\EXCEL.EXE";
    ShellExecute(NULL, "open", executable.c_str(), args.c_str(), directory.c_str(), SW_NORMAL);

    return 0;
}
