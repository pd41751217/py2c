git clone https://github.com/microsoft/vcpkg.git C:\vcpkg
cd C:\vcpkg

.\bootstrap-vcpkg.bat

.\vcpkg integrate install

.\vcpkg install xlnt:x64-mingw-static